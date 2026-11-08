# Blueprint central Outlook signatures

Free, static central template distribution for the Test-User pilot. GitHub Pages hosts the code and generic configuration. Existing Microsoft 365 licensing is required; no paid add-on, database, Graph application, or additional hosting subscription is used.

## Manage the shared template

1. Open `signature/config.json` in this repository and choose **Edit** while signed in with repository write access.
2. Edit `fields.greeting`, `fields.organization`, the HTML/plain-text entries under `templates.standard`, or the `assignments` for `newMail`, `reply`, and `forward`.
3. Change `revision` for every template update, for example from `BP-CENTRAL-V1` to `BP-CENTRAL-V2`. Keep it short and distinctive.
4. Commit to `main`. Wait for the **pages build and deployment** run to succeed. Open the published `signature/config.json` with a fresh query string and confirm the revision.
5. Create a **fresh** compose window in Outlook and verify its signature. Existing drafts and previously sent messages are not changed.

The runtime fetches the shared configuration for each fresh compose event. A template edit does not require a new manifest or per-user setup. Ordinary Pages/CDN publication time still applies.

## Logo

The `logo` object holds a small PNG as base64 plus its display dimensions and alternative text. The runtime attaches it inline and refers to its CID, rather than relying on the recipient downloading a remote logo. Keep it under the validator's limit. The initial pilot image is from Microsoft's MIT-licensed sample.

## Sender information and public content

Outlook provides `displayName` and `emailAddress` locally. They are escaped before insertion and are not included in configuration request URLs, request bodies, telemetry, or the public configuration. Do not add employee directories, credentials, tokens, or confidential material to this public repository.

GitHub receives ordinary web requests and may log request metadata according to its own service policies. This application has no telemetry endpoint and sends no message body, recipient list, or sender profile to a server.

## Pilot boundary

The existing add-in ID is retained and its manifest version is increased. Replace its manifest through Microsoft 365 Integrated apps and keep the assignment to **Test-User only**. The central runtime ignores the original sample's per-user roaming settings.

The configuration contains an additional SHA-256 sender allowlist. It is an accidental-rollout guard, not an authentication or access-control mechanism. Only change it as part of an explicitly authorized rollout. The Microsoft 365 assignment is the primary deployment boundary.

The status pane is read-only. Its GitHub edit link still requires repository write access. No GitHub token or administrative credential is present in the add-in.

## Hosting

Use this account-root Pages repository and publish from `main` / root. Keep `.nojekyll`; it ensures the `.well-known` folder is served. Classic Outlook uses the origin-root file `/.well-known/microsoft-officeaddins-allowed.json`, which lists the exact single-file runtime URL. Use HTTPS.

## Limits and recovery

- Requires an internet connection and supported Outlook add-in APIs. This is not a mandatory send-time signature policy.
- Automatic activation is for fresh messages, replies, and forwards. Existing drafts are outside this first pilot's event scope.
- A user can edit or delete a compose-time signature. Existing native signature behavior may vary; no native defaults are changed by this pilot.
- The pilot checks the initial From address against the Outlook profile. Shared-mailbox and alias sending are outside this pilot. Changing From after insertion is not handled automatically; inspect the signature if you do so.
- Configuration failures preserve the current body/signature and display a warning. Inline-image failures fall back to a signature without the image and show a warning.
- A reply/forward keeps historical signatures in the quote and inserts one fresh signature above it.
- To stop new insertion, set `enabled` to `false` and publish. For emergency rollback, restore the previously tested Microsoft sample manifest in Microsoft 365; manifest synchronization can take time.
- GitHub Pages is a free hosting service for public repositories, with service limits and no bespoke availability guarantee. Review its current service rules before expanding production use.

## License and source

Based on [Microsoft's Outlook signature sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature), licensed under MIT. See `LICENSE`. The runtime and management behavior have been adapted for this pilot.
