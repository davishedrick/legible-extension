const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function makeGoogleDoc(text, revisionId = "revision") {
  return {
    revisionId,
    body: {
      content: [
        {
          paragraph: {
            elements: [
              {
                textRun: {
                  content: `${text}\n`
                }
              }
            ]
          }
        }
      ]
    }
  };
}

function loadBackground(fixtures = [], options = {}) {
  const storage = {};
  const responses = [...fixtures];
  const exports = {};
  const fetchUrls = [];
  const fetchCalls = [];
  const source = fs.readFileSync(path.resolve(__dirname, "..", "background.js"), "utf8");
  const context = {
    console,
    __ACE_TEST_EXPORTS__: exports,
    URLSearchParams,
    fetch: async (url, fetchOptions = {}) => {
      fetchUrls.push(String(url));
      fetchCalls.push({ url: String(url), options: fetchOptions });
      if (!responses.length) {
        if (options.defaultFetchPayload) {
          return {
            ok: true,
            status: 200,
            json: async () => options.defaultFetchPayload,
            text: async () => JSON.stringify(options.defaultFetchPayload)
          };
        }
        throw new Error("No mocked Google Docs response queued.");
      }
      const fixture = responses.shift();
      const payload = fixture && Object.prototype.hasOwnProperty.call(fixture, "payload")
        ? fixture.payload
        : fixture;
      const ok = fixture && Object.prototype.hasOwnProperty.call(fixture, "ok")
        ? Boolean(fixture.ok)
        : true;
      const status = fixture && Object.prototype.hasOwnProperty.call(fixture, "status")
        ? Number(fixture.status)
        : 200;
      return {
        ok,
        status,
        statusText: fixture?.statusText || "",
        json: async () => payload,
        text: async () => JSON.stringify(payload)
      };
    },
    chrome: {
      runtime: {
        lastError: null,
        onMessage: { addListener() {} },
        getManifest() {
          return { oauth2: { client_id: "test-client" } };
        }
      },
      identity: {
        getAuthToken(_options, callback) {
          callback("token");
        },
        removeCachedAuthToken(_options, callback) {
          callback();
        }
      },
      cookies: {
        get(_details, callback) {
          callback(options.sessionCookie ? { value: options.sessionCookie } : null);
        }
      },
      storage: {
        local: {
          get(keys, callback) {
            if (Array.isArray(keys)) {
              callback(Object.fromEntries(keys.map((key) => [key, storage[key]])));
              return;
            }
            if (typeof keys === "string") {
              callback({ [keys]: storage[keys] });
              return;
            }
            callback({ ...storage });
          },
          set(values, callback) {
            Object.assign(storage, values);
            callback();
          },
          remove(keys, callback) {
            (Array.isArray(keys) ? keys : [keys]).forEach((key) => delete storage[key]);
            callback();
          }
        }
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return { exports, storage, fetchUrls, fetchCalls };
}

function loadContent(options = {}) {
  const storage = { ...(options.initialStorage || {}) };
  const exports = {};
  const source = fs.readFileSync(path.resolve(__dirname, "..", "content.js"), "utf8");
  const visibleWordCountElement = options.visibleWordCount === undefined ? null : {
    textContent: `${Number(options.visibleWordCount) || 0} words`,
    innerText: `${Number(options.visibleWordCount) || 0} words`,
    parentElement: null,
    closest() { return null; },
    getAttribute() { return ""; },
    getBoundingClientRect() {
      return { left: 20, top: 690, right: 260, bottom: 744, width: 240, height: 54 };
    }
  };
  const fakeElement = () => ({
    setAttribute() {},
    appendChild() {},
    addEventListener() {},
    querySelector() { return null; },
    contains() { return false; },
    getBoundingClientRect() {
      return { left: 0, top: 0, width: 0, height: 0 };
    },
    style: {},
    remove() {}
  });
  const location = {
    href: "https://docs.google.com/document/d/doc-test/edit",
    pathname: "/document/d/doc-test/edit",
    search: "",
    hash: "",
    ...(options.location || {})
  };
  const window = {
    addEventListener() {},
    requestAnimationFrame(callback) { callback(); },
    setInterval(callback, milliseconds) {
      const handle = setInterval(callback, milliseconds);
      handle.unref?.();
      return handle;
    },
    clearInterval,
    setTimeout,
    clearTimeout,
    innerWidth: 1024,
    innerHeight: 768,
    frames: [],
    location
  };
  window.top = options.topFrame ? window : {};
  const document = {
    activeElement: null,
    documentElement: { ...fakeElement(), clientWidth: 1024, clientHeight: 768 },
    getElementById() { return null; },
    createElement: fakeElement,
    addEventListener() {},
    querySelectorAll() {
      return visibleWordCountElement ? [visibleWordCountElement] : [];
    },
    elementsFromPoint() {
      return visibleWordCountElement ? [visibleWordCountElement] : [];
    }
  };
  window.document = document;
  const context = {
    console,
    __ACE_TEST_EXPORTS__: exports,
    window,
    document,
    navigator: { clipboard: null },
    sessionStorage: {
      getItem() { return null; },
      setItem() {}
    },
    chrome: {
      runtime: {
        lastError: null,
        sendMessage(message, callback) {
          if (message.aceType === "ace-current-tab-scope") {
            if (callback) {
              callback({
                ok: true,
                chromeTabId: options.chromeTabId ?? 1,
                frameId: 0
              });
            }
            return;
          }
          if (options.sendMessage) {
            options.sendMessage(message, callback);
            return;
          }
          if (callback) {
            callback({
              ok: false,
              wordCount: null,
              error: "No mocked runtime response queued."
            });
          }
        }
      },
      storage: {
        local: {
          get(keys, callback) {
            if (Array.isArray(keys)) {
              callback(Object.fromEntries(keys.map((key) => [key, storage[key]])));
              return;
            }
            if (typeof keys === "string") {
              callback({ [keys]: storage[keys] });
              return;
            }
            callback({ ...storage });
          },
          set(values, callback) {
            Object.assign(storage, values);
            callback();
          },
          remove(keys, callback) {
            (Array.isArray(keys) ? keys : [keys]).forEach((key) => delete storage[key]);
            callback();
          }
        }
      }
    },
    location,
    setTimeout,
    clearTimeout,
    URLSearchParams
  };
  vm.createContext(context);
  vm.runInContext(source, context);
  return { exports, storage, context };
}

module.exports = {
  loadBackground,
  loadContent,
  makeGoogleDoc
};
