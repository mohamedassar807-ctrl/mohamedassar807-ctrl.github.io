/* Blueprint pilot: read-only preview and a local, explicit draft-marker check. */
/* global Office, BlueprintSignature */
(function () {
  "use strict";
  var config;
  var profile;
  var configUrl = "https://mohamedassar807-ctrl.github.io/signature/config.json";
  function el(id) { return document.getElementById(id); }
  var runtimeVersion = typeof BlueprintSignature.runtimeVersion === "string" ? BlueprintSignature.runtimeVersion : "unavailable";
  var supportsV2 = typeof BlueprintSignature.supportsSchemaVersion === "function" && BlueprintSignature.supportsSchemaVersion(2);
  el("runtime-info").textContent = "Runtime " + runtimeVersion + " · policy schema 2 " + (supportsV2 ? "supported" : "not supported");
  function checkDraft() {
    var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    var finished = false;
    el("check-draft").disabled = true;
    el("draft-check").textContent = "Checking this draft locally…";
    var timer = setTimeout(function () { finish("Unconfirmed. The draft check timed out; try again in a fresh new HTML message."); }, 5000);
    function finish(message) {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      el("draft-check").textContent = message;
      el("check-draft").disabled = false;
    }
    function invoke(owner, method, args, callback) {
      if (finished) return;
      if (!owner || typeof owner[method] !== "function") { finish("Unconfirmed. This draft does not expose the required Outlook APIs."); return; }
      var responded = false;
      try {
        owner[method].apply(owner, args.concat([function (result) {
          if (finished || responded) return;
          responded = true;
          if (!result || !(result.status === "succeeded" || (Office.AsyncResultStatus && result.status === Office.AsyncResultStatus.Succeeded))) {
            finish("Unconfirmed. Outlook could not read this draft for the upgrade check.");
            return;
          }
          try { callback(result.value); }
          catch (error) { finish("Unconfirmed. Outlook could not complete this draft check."); }
        }]));
      } catch (error) { finish("Unconfirmed. Outlook could not complete this draft check."); }
    }
    if (!supportsV2 || typeof BlueprintSignature.htmlRuntimeMarker !== "string") {
      finish("Unconfirmed. This status pane did not load the V2-capable runtime.");
      return;
    }
    invoke(item, "getComposeTypeAsync", [], function (compose) {
      if (!compose || (compose.composeType !== "newEmail" && compose.composeType !== "newMail")) {
        finish("Unconfirmed. Use a fresh new HTML message; a marker in quoted reply or forward history cannot confirm this draft.");
        return;
      }
      invoke(item.body, "getTypeAsync", [], function (bodyType) {
        if (bodyType !== "html") {
          finish("Unconfirmed. Compose a fresh new HTML message for this upgrade check; plain text has no hidden runtime marker.");
          return;
        }
        var coercion = Office.CoercionType ? Office.CoercionType.Html : "html";
        invoke(item.body, "getAsync", [coercion], function (html) {
          // The body stays inside this callback. It is not logged, persisted,
          // sent in a request, or returned through a host bridge.
          finish(typeof html === "string" && html.indexOf(BlueprintSignature.htmlRuntimeMarker) !== -1
            ? "Runtime " + runtimeVersion + " applied this draft."
            : "Unconfirmed. The expected runtime marker was not found. Check a fresh new HTML message after the add-in update.");
        });
      });
    });
  }
  function preview() {
    if (!config || !profile) return;
    var html = BlueprintSignature.render(config, profile, el("compose-type").value, "preview-logo.png", "html");
    html = html.replace('cid:preview-logo.png', 'data:image/png;base64,' + config.logo.base64);
    el("preview").srcdoc = '<!doctype html><html><head><meta charset="utf-8"><meta name="referrer" content="no-referrer"></head><body>' + html + '</body></html>';
  }
  function load() {
    el("refresh").disabled = true;
    el("status").textContent = "Loading the shared template…";
    var request = new XMLHttpRequest();
    request.open("GET", configUrl + "?refresh=" + Date.now(), true);
    request.timeout = 5000;
    request.onload = function () {
      try {
        if (request.status !== 200 || request.responseText.length > 131072) throw new Error("Template is not available.");
        config = BlueprintSignature.validateConfig(JSON.parse(request.responseText));
        var allowed = config.pilot.allowedSenderSha256.indexOf(BlueprintSignature.sha256(profile.emailAddress.trim().toLowerCase())) !== -1;
        el("revision").textContent = config.revision;
        el("profile").textContent = profile.displayName + " · " + profile.emailAddress;
        el("scope").textContent = allowed ? "This profile is included. Automatic insertion also checks the From address." : "This profile is outside the pilot.";
        el("details").hidden = false;
        el("status").textContent = config.enabled && allowed ? "The shared template is available for this profile. No per-user setup is needed." : "Automatic insertion is not active for this profile.";
        preview();
      } catch (error) {
        config = null;
        el("preview").srcdoc = "";
        el("details").hidden = true;
        el("status").textContent = "The shared template could not be loaded or validated. Your current message has not been changed.";
      }
      el("refresh").disabled = false;
    };
    request.onerror = request.ontimeout = function () {
      config = null;
      el("preview").srcdoc = "";
      el("details").hidden = true;
      el("status").textContent = "The shared template is unreachable. Your current message has not been changed.";
      el("refresh").disabled = false;
    };
    request.send();
  }
  Office.onReady(function () {
    profile = Office.context.mailbox.userProfile;
    el("compose-type").addEventListener("change", preview);
    el("refresh").addEventListener("click", load);
    el("check-draft").addEventListener("click", checkDraft);
    el("check-draft").disabled = false;
    load();
  });
})();
