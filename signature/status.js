/* Blueprint pilot: preview only. No message body or recipient data is read here. */
/* global Office, BlueprintSignature */
(function () {
  "use strict";
  var config;
  var profile;
  var configUrl = "https://mohamedassar807-ctrl.github.io/signature/config.json";
  function el(id) { return document.getElementById(id); }
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
    load();
  });
})();
