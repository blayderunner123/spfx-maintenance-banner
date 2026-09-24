# Provenance

- **Recovered directly:** IDs, versions, package metadata, permission request, feature XML, localization output, bundled icon, bundle name, selected fields, query shape, and runtime behavior visible in the production bundle.
- **Reconstructed:** readable TypeScript types, method names/structure, imports, and formatting that reproduce the compiled behavior.
- **Recreated scaffolding:** npm/build configuration, TypeScript configuration, development serve configuration, and dependency manifest needed for an SPFx 1.20 build.

The npm manifest contains a narrow override for an invalid transitive dependency published in `@microsoft/spfx-heft-plugins@1.20.0` (`@microsoft/sp-module-interfaces@1.19.0-dev.233`, no longer available from npm). It resolves that dependency to `1.20.2`, the same stable interface version required by Microsoft's SPFx 1.20 runtime packages. This is recreated build scaffolding, not recovered application behavior.

The compiled bundle includes PnPjs 4.19.0 and targets a fixed organization-specific configuration web. The public repository intentionally replaces that tenant URL with the optional `configurationSiteUrl` Application Customizer property and defaults to the current web. This is a documented publication hardening change from the recovered production behavior.

The artifact does not contain a list provisioning schema. A production List Settings screenshot supplied during recovery directly confirms the following display names and field types (only `Title` is visibly marked required):

| Field | Confirmed type |
|---|---|
| Title | Single line of text (required) |
| Message | Multiple lines of text |
| Enabled | Yes/No |
| BannerType | Choice |
| StartDate | Date and Time |
| EndDate | Date and Time |
| TargetSites | Multiple lines of text |
| ExcludedSites | Multiple lines of text |
| Priority | Number |
| LinkText | Single line of text |
| LinkUrl | Hyperlink or Picture |
| Dismissible | Yes/No |
| TargetUsers | Multiple lines of text |
| TargetAudience | Person or Group |
| ShowCountdown | Yes/No |
| CountdownLabel | Single line of text |
| OpenLinkInNewTab | Yes/No |
| AutoRefreshSeconds | Number |

The screenshot also shows the standard `Modified`, `Created`, `Created By`, and `Modified By` fields. It does not establish the `BannerType` choice values, default values, date display formats, multiline text modes, numeric bounds, or whether `TargetAudience` allows multiple selections. Runtime expansion of `TargetAudience` as an array strongly indicates multiple values, but that remains a compiled-code inference. No unsupported provisioning feature was invented.
