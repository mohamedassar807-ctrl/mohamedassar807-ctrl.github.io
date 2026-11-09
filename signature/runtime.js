/* Blueprint central signatures. Based on the Microsoft Outlook signature sample.
 * Copyright (c) Microsoft Corporation. Licensed under the MIT License.
 * This file deliberately contains the complete Classic Outlook event runtime.
 */
(function (root) {
  "use strict";

  var CONFIG_URL = "https://mohamedassar807-ctrl.github.io/signature/config.json";
  var RUNTIME_VERSION = "1.2.0.0";
  var HTML_RUNTIME_MARKER = "<!--BlueprintRuntime:" + RUNTIME_VERSION + "-->";
  var MAX_CONFIG_BYTES = 128 * 1024;
  var MAX_LOGO_BASE64 = 64 * 1024;
  var notificationKey = "blueprint-central-signature-status";
  var cidCounter = 0;
  var own = Object.prototype.hasOwnProperty;
  var tokenNames = { greeting: 1, displayName: 1, email: 1, organization: 1, revision: 1, logo: 1 };

  function fail(message) { throw new Error(message); }
  function isObject(value) { return value !== null && typeof value === "object" && !Array.isArray(value); }
  function checkKeys(value, allowed, label) {
    if (!isObject(value)) fail(label + " must be an object.");
    Object.keys(value).forEach(function (key) {
      if (allowed.indexOf(key) === -1) fail("Unknown " + label + " property: " + key);
    });
  }
  function string(value, limit, label, allowEmpty) {
    if (typeof value !== "string" || value.length > limit || (!allowEmpty && !value.length) || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value)) {
      fail("Invalid " + label + ".");
    }
    return value;
  }
  function utf8Bytes(text) {
    var bytes = [];
    for (var i = 0; i < text.length; i += 1) {
      var point = text.charCodeAt(i);
      if (point >= 0xd800 && point <= 0xdbff) {
        var low = text.charCodeAt(i + 1);
        if (low >= 0xdc00 && low <= 0xdfff) { point = 0x10000 + ((point - 0xd800) << 10) + low - 0xdc00; i += 1; }
        else point = 0xfffd;
      } else if (point >= 0xdc00 && point <= 0xdfff) point = 0xfffd;
      if (point < 0x80) bytes.push(point);
      else if (point < 0x800) bytes.push(0xc0 | (point >> 6), 0x80 | (point & 63));
      else if (point < 0x10000) bytes.push(0xe0 | (point >> 12), 0x80 | ((point >> 6) & 63), 0x80 | (point & 63));
      else bytes.push(0xf0 | (point >> 18), 0x80 | ((point >> 12) & 63), 0x80 | ((point >> 6) & 63), 0x80 | (point & 63));
    }
    return bytes;
  }

  // Synchronous SHA-256 avoids assuming WebCrypto exists in Classic's JS runtime.
  function sha256(text) {
    if (typeof text !== "string") fail("SHA-256 input must be text.");
    var bytes = utf8Bytes(text);
    var length = bytes.length;
    bytes.push(0x80);
    while (bytes.length % 64 !== 56) bytes.push(0);
    var high = Math.floor(length / 0x20000000);
    var low = (length << 3) >>> 0;
    bytes.push((high >>> 24) & 255, (high >>> 16) & 255, (high >>> 8) & 255, high & 255,
      (low >>> 24) & 255, (low >>> 16) & 255, (low >>> 8) & 255, low & 255);
    var h = [0x6a09e667, 0xbb67ae85, 0x3c6ef372, 0xa54ff53a, 0x510e527f, 0x9b05688c, 0x1f83d9ab, 0x5be0cd19];
    var k = [0x428a2f98,0x71374491,0xb5c0fbcf,0xe9b5dba5,0x3956c25b,0x59f111f1,0x923f82a4,0xab1c5ed5,
      0xd807aa98,0x12835b01,0x243185be,0x550c7dc3,0x72be5d74,0x80deb1fe,0x9bdc06a7,0xc19bf174,
      0xe49b69c1,0xefbe4786,0x0fc19dc6,0x240ca1cc,0x2de92c6f,0x4a7484aa,0x5cb0a9dc,0x76f988da,
      0x983e5152,0xa831c66d,0xb00327c8,0xbf597fc7,0xc6e00bf3,0xd5a79147,0x06ca6351,0x14292967,
      0x27b70a85,0x2e1b2138,0x4d2c6dfc,0x53380d13,0x650a7354,0x766a0abb,0x81c2c92e,0x92722c85,
      0xa2bfe8a1,0xa81a664b,0xc24b8b70,0xc76c51a3,0xd192e819,0xd6990624,0xf40e3585,0x106aa070,
      0x19a4c116,0x1e376c08,0x2748774c,0x34b0bcb5,0x391c0cb3,0x4ed8aa4a,0x5b9cca4f,0x682e6ff3,
      0x748f82ee,0x78a5636f,0x84c87814,0x8cc70208,0x90befffa,0xa4506ceb,0xbef9a3f7,0xc67178f2];
    function rotr(x, n) { return (x >>> n) | (x << (32 - n)); }
    var w = new Array(64);
    for (var offset = 0; offset < bytes.length; offset += 64) {
      for (var j = 0; j < 16; j += 1) {
        var at = offset + j * 4;
        w[j] = ((bytes[at] << 24) | (bytes[at + 1] << 16) | (bytes[at + 2] << 8) | bytes[at + 3]) >>> 0;
      }
      for (j = 16; j < 64; j += 1) {
        var x = w[j - 15]; var y = w[j - 2];
        var s0 = rotr(x, 7) ^ rotr(x, 18) ^ (x >>> 3);
        var s1 = rotr(y, 17) ^ rotr(y, 19) ^ (y >>> 10);
        w[j] = (w[j - 16] + s0 + w[j - 7] + s1) >>> 0;
      }
      var a=h[0], b=h[1], c=h[2], d=h[3], e=h[4], f=h[5], g=h[6], v=h[7];
      for (j = 0; j < 64; j += 1) {
        var sum1 = rotr(e,6) ^ rotr(e,11) ^ rotr(e,25);
        var choose = (e & f) ^ (~e & g);
        var t1 = (v + sum1 + choose + k[j] + w[j]) >>> 0;
        var sum0 = rotr(a,2) ^ rotr(a,13) ^ rotr(a,22);
        var majority = (a & b) ^ (a & c) ^ (b & c);
        var t2 = (sum0 + majority) >>> 0;
        v=g; g=f; f=e; e=(d+t1)>>>0; d=c; c=b; b=a; a=(t1+t2)>>>0;
      }
      h[0]=(h[0]+a)>>>0; h[1]=(h[1]+b)>>>0; h[2]=(h[2]+c)>>>0; h[3]=(h[3]+d)>>>0;
      h[4]=(h[4]+e)>>>0; h[5]=(h[5]+f)>>>0; h[6]=(h[6]+g)>>>0; h[7]=(h[7]+v)>>>0;
    }
    return h.map(function (value) { return ("00000000" + value.toString(16)).slice(-8); }).join("");
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, function (ch) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[ch];
    });
  }
  function normalizeProfile(profile) {
    if (!profile) fail("Mailbox profile is unavailable.");
    var email = string(profile.emailAddress, 320, "mailbox email", false).trim().toLowerCase();
    if (!/^[a-z0-9.!#$%&'*+\/=^_`{|}~-]+@[a-z0-9](?:[a-z0-9.-]*[a-z0-9])?$/i.test(email)) fail("Invalid mailbox email.");
    var displayName = string(profile.displayName, 256, "display name", true);
    return { displayName: displayName || email, emailAddress: email };
  }
  function normalizeComposeType(type) {
    if (type === "newEmail" || type === "newMail") return "newMail";
    if (type === "reply" || type === "forward") return type;
    fail("Unsupported compose type.");
  }
  function checkTokens(value, label) {
    var remainder = value.replace(/\{\{([^{}]+)\}\}/g, function (whole, name) {
      if (!own.call(tokenNames, name)) fail("Unknown " + label + " token.");
      return "";
    });
    if (/\{\{|\}\}/.test(remainder)) fail("Malformed " + label + " token.");
  }

  var styleProperties = ("font-family font-size font-weight font-style line-height color background-color border border-top border-right border-bottom border-left border-collapse padding padding-top padding-right padding-bottom padding-left margin margin-top margin-right margin-bottom margin-left text-decoration vertical-align width height display text-align white-space").split(" ");
  function checkStyle(value) {
    if (/[{}@\\<>]|url\s*\(|expression\s*\(|javascript|vbscript|data\s*:/i.test(value)) fail("Unsafe template style.");
    value.split(";").forEach(function (declaration) {
      if (!declaration.trim()) return;
      var pair = declaration.split(":");
      if (pair.length !== 2 || styleProperties.indexOf(pair[0].trim().toLowerCase()) === -1 || !/^[a-z0-9#.,% \t'"/_-]+$/i.test(pair[1].trim())) fail("Unsupported template style.");
    });
  }
  function validateHtml(html) {
    var tags = ["div","p","br","strong","b","em","i","span","table","tbody","tr","td","a"];
    var tagPattern = /<[^>]*>/g;
    var previous = 0;
    var stack = [];
    var match;
    while ((match = tagPattern.exec(html)) !== null) {
      if (/[<>]/.test(html.slice(previous, match.index))) fail("Malformed template markup.");
      previous = tagPattern.lastIndex;
      var tag = /^<(\/?)\s*([a-z][a-z0-9]*)([^<>]*?)(\/?)>$/i.exec(match[0]);
      if (!tag) fail("Unsupported template markup.");
      var name = tag[2].toLowerCase();
      if (tags.indexOf(name) === -1) fail("Unsupported template tag: " + name);
      if (tag[1]) {
        if (tag[3].trim() || tag[4] || stack.pop() !== name) fail("Unbalanced template markup.");
        continue;
      }
      var attrs = tag[3];
      var attributePattern = /\s+([a-z][a-z0-9-]*)\s*=\s*("[^"]*"|'[^']*')/ig;
      var attrMatch; var attrEnd = 0; var seen = {};
      while ((attrMatch = attributePattern.exec(attrs)) !== null) {
        if (attrs.slice(attrEnd, attrMatch.index).trim()) fail("Malformed template attribute.");
        attrEnd = attributePattern.lastIndex;
        var attr = attrMatch[1].toLowerCase(); var value = attrMatch[2].slice(1, -1);
        if (own.call(seen, attr)) fail("Duplicate template attribute.");
        seen[attr] = true;
        if (attr === "style") checkStyle(value);
        else if (name === "a" && attr === "href" && value === "mailto:{{email}}") { /* fixed local identity substitution */ }
        else if (name === "table" && attr === "role" && value === "presentation") { /* accessible layout table */ }
        else if (name === "table" && /^(cellpadding|cellspacing|border)$/.test(attr) && /^\d{1,2}$/.test(value)) { /* numeric layout */ }
        else if ((name === "table" || name === "td") && attr === "width" && /^\d{1,4}%?$/.test(value)) { /* numeric layout */ }
        else if (name === "td" && /^(colspan|rowspan)$/.test(attr) && /^[1-9]\d?$/.test(value)) { /* numeric layout */ }
        else if (name === "td" && attr === "valign" && /^(top|middle|bottom)$/.test(value)) { /* numeric layout */ }
        else if (attr === "align" && /^(left|center|right)$/.test(value)) { /* layout */ }
        else fail("Unsupported template attribute: " + attr);
      }
      if (attrs.slice(attrEnd).trim()) fail("Malformed template attribute.");
      if (name !== "br" && !tag[4]) stack.push(name);
      else if (name !== "br") fail("Only br may be self-closing.");
    }
    if (/[<>]/.test(html.slice(previous)) || stack.length) fail("Unbalanced template markup.");
  }
  function validatePng(base64) {
    if (typeof base64 !== "string" || base64.length < 12 || base64.length > MAX_LOGO_BASE64 || base64.length % 4 || !/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) fail("Invalid PNG base64.");
    var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    var prefix = [];
    for (var i = 0; i < 12; i += 4) {
      var packed = (alphabet.indexOf(base64.charAt(i)) << 18) | (alphabet.indexOf(base64.charAt(i + 1)) << 12) | (alphabet.indexOf(base64.charAt(i + 2)) << 6) | alphabet.indexOf(base64.charAt(i + 3));
      prefix.push((packed >>> 16) & 255, (packed >>> 8) & 255, packed & 255);
    }
    var signature = [137,80,78,71,13,10,26,10];
    for (i = 0; i < signature.length; i += 1) if (prefix[i] !== signature[i]) fail("Logo is not PNG data.");
  }
  function supportsSchemaVersion(version) { return version === 1 || version === 2; }
  function validateConfig(config) {
    var serialized;
    try { serialized = JSON.stringify(config); } catch (error) { fail("Configuration cannot be serialized."); }
    if (typeof serialized !== "string" || utf8Bytes(serialized).length > MAX_CONFIG_BYTES) fail("Configuration exceeds 128 KiB.");
    var configKeys = ["schemaVersion","revision","enabled","pilot","fields","logo","assignments","templates"];
    if (config && config.schemaVersion === 2) configKeys.push("userAssignments");
    checkKeys(config, configKeys, "configuration");
    if (!supportsSchemaVersion(config.schemaVersion) || typeof config.enabled !== "boolean") fail("Unsupported configuration schema.");
    if (!/^[a-z0-9][a-z0-9._-]{0,63}$/i.test(string(config.revision, 64, "revision", false))) fail("Invalid revision.");
    checkKeys(config.pilot, ["allowedSenderSha256"], "pilot");
    var allowed = config.pilot.allowedSenderSha256;
    if (!Array.isArray(allowed) || (config.schemaVersion === 1 && !allowed.length) || allowed.length > 100 || allowed.some(function (hash, index) { return typeof hash !== "string" || !/^[a-f0-9]{64}$/.test(hash) || (config.schemaVersion === 2 && allowed.indexOf(hash) !== index); })) fail("Invalid pilot allow-list.");
    checkKeys(config.fields, ["greeting","organization"], "fields");
    string(config.fields.greeting, 500, "greeting", true); string(config.fields.organization, 1000, "organization", true);
    checkKeys(config.logo, ["base64","fileName","alt","width","height"], "logo");
    validatePng(config.logo.base64);
    if (!/^[a-z0-9][a-z0-9._-]{0,63}\.png$/i.test(config.logo.fileName) || config.logo.fileName.indexOf("..") !== -1) fail("Invalid logo filename.");
    string(config.logo.alt, 160, "logo alt text", true);
    [config.logo.width, config.logo.height].forEach(function (dimension) {
      if (typeof dimension !== "number" || dimension % 1 || dimension < 1 || dimension > 256) fail("Invalid logo dimensions.");
    });
    if (!isObject(config.templates) || !Object.keys(config.templates).length || Object.keys(config.templates).length > 10) fail("Invalid templates.");
    Object.keys(config.templates).forEach(function (name) {
      if (!/^[a-z0-9_-]{1,40}$/i.test(name)) fail("Invalid template name.");
      var template = config.templates[name];
      checkKeys(template, ["html","text"], "template");
      string(template.html, 20000, "HTML template", false); string(template.text, 10000, "text template", false);
      checkTokens(template.html, "HTML template"); checkTokens(template.text, "text template");
      validateHtml(template.html);
    });
    function validateAssignments(assignments) {
      checkKeys(assignments, ["newMail","reply","forward"], "assignments");
      ["newMail","reply","forward"].forEach(function (composeType) {
        var name = assignments[composeType];
        if (!own.call(assignments, composeType) || (name !== "none" && (typeof name !== "string" || !own.call(config.templates, name)))) fail("Invalid template assignment.");
      });
    }
    validateAssignments(config.assignments);
    if (config.schemaVersion === 2) {
      if (!isObject(config.userAssignments) || Object.keys(config.userAssignments).length > 100) fail("Invalid user assignments map.");
      Object.keys(config.userAssignments).forEach(function (hash) {
        if (!/^[a-f0-9]{64}$/.test(hash)) fail("Invalid user assignments hash.");
        validateAssignments(config.userAssignments[hash]);
      });
    }
    return config;
  }
  function assignmentsForProfile(config, normalizedProfile) {
    var hash = sha256(normalizedProfile.emailAddress);
    return config.schemaVersion === 2 && own.call(config.userAssignments, hash) ? config.userAssignments[hash] : config.assignments;
  }
  function resolveAssignments(config, profile) {
    validateConfig(config);
    var assignments = assignmentsForProfile(config, normalizeProfile(profile));
    return { newMail: assignments.newMail, reply: assignments.reply, forward: assignments.forward };
  }
  function isAllowedSender(config, profile) {
    var normalized = normalizeProfile(profile);
    return config.pilot.allowedSenderSha256.indexOf(sha256(normalized.emailAddress)) !== -1;
  }
  function render(config, profile, composeType, logoCid, bodyType) {
    validateConfig(config);
    var normalized = normalizeProfile(profile);
    var type = normalizeComposeType(composeType);
    if (bodyType !== "html" && bodyType !== "text") fail("Unsupported body format.");
    var assignment = assignmentsForProfile(config, normalized)[type];
    if (!config.enabled || assignment === "none") return "";
    var values = {
      greeting: config.fields.greeting,
      displayName: normalized.displayName,
      email: normalized.emailAddress,
      organization: config.fields.organization,
      revision: config.revision,
      logo: ""
    };
    var template = config.templates[assignment][bodyType];
    if (bodyType === "html") {
      // This is the only URI-bearing template attribute. Encode the address before
      // generic text escaping so SMTP punctuation cannot create mailto parameters.
      template = template.replace(/href\s*=\s*(["'])mailto:\{\{email\}\}\1/gi, function () {
        return 'href="mailto:' + escapeHtml(encodeURIComponent(normalized.emailAddress).replace(/%40/gi, "@")) + '"';
      });
      Object.keys(values).forEach(function (name) { if (name !== "logo") values[name] = escapeHtml(values[name]); });
      if (logoCid) {
        if (typeof logoCid !== "string" || !/^[a-z0-9][a-z0-9._-]{0,159}\.png$/i.test(logoCid)) fail("Invalid inline logo CID.");
        values.logo = '<img src="cid:' + logoCid + '" alt="' + escapeHtml(config.logo.alt) + '" width="' + config.logo.width + '" height="' + config.logo.height + '">';
      }
    }
    var result = template.replace(/\{\{([^{}]+)\}\}/g, function (whole, name) { return values[name]; });
    if (result.length > 30000) fail("Rendered signature exceeds Outlook's limit.");
    return result;
  }

  function createHandler(deps) {
    deps = deps || {};
    var office = deps.office || root.Office;
    var fetchFn = deps.fetch || (typeof root.fetch === "function" ? root.fetch.bind(root) : null);
    var setTimer = deps.setTimeout || root.setTimeout;
    var clearTimer = deps.clearTimeout || root.clearTimeout;
    var now = deps.now || Date.now;
    var random = deps.random || Math.random;
    var fetchTimeoutMs = deps.fetchTimeoutMs || 5000;
    var apiTimeoutMs = deps.apiTimeoutMs || 5000;

    function timed(work, milliseconds, label) {
      return new Promise(function (resolve, reject) {
        var settled = false;
        var timer = setTimer(function () {
          if (!settled) {
            settled = true;
            var timeoutError = new Error(label + " timed out.");
            timeoutError.code = "timeout";
            reject(timeoutError);
          }
        }, milliseconds);
        Promise.resolve(work).then(function (value) {
          if (!settled) { settled = true; clearTimer(timer); resolve(value); }
        }, function (error) {
          if (!settled) { settled = true; clearTimer(timer); reject(error); }
        });
      });
    }
    function invoke(owner, method, args) {
      return timed(new Promise(function (resolve, reject) {
        if (!owner || typeof owner[method] !== "function") { reject(new Error(method + " is unavailable.")); return; }
        try {
          owner[method].apply(owner, args.concat([function (result) {
            if (result && (result.status === "succeeded" || (office.AsyncResultStatus && result.status === office.AsyncResultStatus.Succeeded))) resolve(result.value);
            else reject(new Error(method + " failed."));
          }]));
        } catch (error) { reject(error); }
      }), apiTimeoutMs, method);
    }
    function notify(item, message) {
      if (!item || !item.notificationMessages || typeof item.notificationMessages.addAsync !== "function") return Promise.resolve();
      return invoke(item.notificationMessages, "addAsync", [notificationKey, { type: "errorMessage", message: message }]).catch(function () { /* notification cannot prevent completion */ });
    }
    function loadConfig() {
      if (!fetchFn) return Promise.reject(new Error("Fetch is unavailable."));
      var url = CONFIG_URL + "?_ts=" + encodeURIComponent(String(now()));
      var work = Promise.resolve().then(function () {
        return fetchFn(url, { method: "GET", credentials: "omit", cache: "no-store", referrerPolicy: "no-referrer" });
      }).then(function (response) {
        if (!response || response.ok !== true || typeof response.text !== "function") fail("Configuration request failed.");
        return response.text();
      }).then(function (text) {
        if (typeof text !== "string" || text.length > MAX_CONFIG_BYTES || utf8Bytes(text).length > MAX_CONFIG_BYTES) fail("Configuration exceeds 128 KiB.");
        return validateConfig(JSON.parse(text));
      });
      return timed(work, fetchTimeoutMs, "Configuration request");
    }
    function uniqueLogoName(config) {
      cidCounter += 1;
      return config.logo.fileName.replace(/\.png$/i, "").slice(0, 32) + "-" + config.revision.slice(0, 24) + "-" + Math.floor(now()).toString(36) + "-" + String(random().toString(36).slice(2, 10) || "0") + "-" + cidCounter.toString(36) + ".png";
    }
    return function (event) {
      var completed = false;
      var mutationStarted = false;
      var item = office && office.context && office.context.mailbox && office.context.mailbox.item;
      function finish() {
        if (completed) return;
        completed = true;
        if (event && typeof event.completed === "function") {
          try { event.completed(); } catch (error) { /* Office owns the lifecycle; never complete twice */ }
        }
      }
      var profile;
      var run = Promise.resolve().then(function () {
        if (!item || !office.context.mailbox.userProfile) fail("Mailbox context is unavailable.");
        profile = normalizeProfile(office.context.mailbox.userProfile);
        return loadConfig();
      }).then(function (config) {
        if (!config.enabled || !isAllowedSender(config, profile)) return;
        // Deployment scope and userProfile guard the mailbox. Also verify the
        // initial From value so an existing delegated/send-as compose is skipped.
        return invoke(item.from, "getAsync", []).then(function (sender) {
          var normalizedSender = normalizeProfile({ displayName: (sender && sender.displayName) || "", emailAddress: sender && sender.emailAddress });
          if (normalizedSender.emailAddress !== profile.emailAddress || !isAllowedSender(config, normalizedSender)) fail("The current sender is outside the pilot.");
          return invoke(item, "getComposeTypeAsync", []);
        }).then(function (compose) {
          var type = normalizeComposeType(compose && compose.composeType);
          if (assignmentsForProfile(config, profile)[type] === "none") return;
          return invoke(item.body, "getTypeAsync", []).then(function (bodyType) {
            if (bodyType !== "html" && bodyType !== "text") fail("Unsupported body format.");
            function insert(logoCid) {
              var signature = render(config, profile, type, logoCid, bodyType);
              if (bodyType === "html") signature = HTML_RUNTIME_MARKER + signature;
              if (signature.length > 30000) fail("Signature payload exceeds Outlook's limit.");
              var coercionType = bodyType;
              if (office.CoercionType) coercionType = bodyType === "html" ? office.CoercionType.Html : office.CoercionType.Text;
              mutationStarted = true;
              return invoke(item.body, "setSignatureAsync", [signature, { coercionType: coercionType }]);
            }
            if (bodyType === "text") return insert(null);
            var filename = uniqueLogoName(config);
            mutationStarted = true;
            return invoke(item, "addFileAttachmentFromBase64Async", [config.logo.base64, filename, { isInline: true }]).then(function () {
              // A failed or delayed signature callback does not prove Outlook
              // left the body untouched. Keep its CID attachment to avoid
              // breaking an insertion which completes after our own timeout.
              return insert(filename);
            }, function (error) {
              if (error && error.code === "timeout") throw error;
              return insert(null).then(function () {
                return notify(item, "Blueprint signature: the logo could not be embedded. A signature without the image was inserted.");
              });
            });
          });
        });
      }).catch(function () {
        return notify(item, mutationStarted
          ? "Blueprint could not confirm the signature update. Inspect the message and logo before sending."
          : "Blueprint signature could not be updated. Existing message content was preserved; check your signature before sending.");
      });
      return run.then(finish, finish);
    };
  }

  function checkSignature(event) { return createHandler()(event); }
  var api = {
    runtimeVersion: RUNTIME_VERSION,
    htmlRuntimeMarker: HTML_RUNTIME_MARKER,
    supportsSchemaVersion: supportsSchemaVersion,
    validateConfig: validateConfig,
    render: render,
    sha256: sha256,
    escapeHtml: escapeHtml,
    normalizeProfile: normalizeProfile,
    normalizeComposeType: normalizeComposeType,
    resolveAssignments: resolveAssignments,
    isAllowedSender: isAllowedSender,
    createHandler: createHandler,
    checkSignature: checkSignature,
    configUrl: CONFIG_URL
  };
  root.BlueprintSignature = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  if (root.Office && root.Office.actions && typeof root.Office.actions.associate === "function") root.Office.actions.associate("checkSignature", checkSignature);
}(typeof globalThis !== "undefined" ? globalThis : (typeof window !== "undefined" ? window : this)));
