# Artifact provenance

This repository is a forensic source reconstruction from the preserved production package `sharepoint-maintenance-banner.sppkg`.

- Golden artifact size: 46,125 bytes
- Golden SHA-256: `56887B997C57DE2459D8321D7CC4C7BF537EFC47357A7367D9A0505319A8822B`
- Solution/component version: 1.3.3
- SPFx runtime version: 1.20.0

The deployed bundle, manifests, package metadata, permissions, feature definitions, icon, and runtime behavior were recovered directly. Readable TypeScript and development scaffolding were reconstructed and are clearly identified in `FORENSIC-NOTES.md`.

The repository's 1.3.4 source is a publication-ready successor that removes organization-specific metadata and replaces the recovered fixed configuration URL with the documented `configurationSiteUrl` property.

The original SPPKG remains the immutable reference artifact and is not stored in this repository.
