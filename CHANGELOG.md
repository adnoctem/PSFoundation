## [1.8.2](https://github.com/adnoctem/PSFoundation/compare/v1.8.1...v1.8.2) (2026-09-30)

### Bug Fixes

* **src:** write module state under PSFoundation, not winkit ([15d5c66](https://github.com/adnoctem/PSFoundation/commit/15d5c660834a97dd93bd25cb1c8a3ea2085b5d9e))

## [1.8.1](https://github.com/adnoctem/PSFoundation/compare/v1.8.0...v1.8.1) (2026-09-29)

### Bug Fixes

- **src:** derive Office builds from active product resource versions
  ([ce9ff77](https://github.com/adnoctem/PSFoundation/commit/ce9ff7779db4a91fe486840ee0fbb1aa8c7309ab))
- **src:** retire Office pilot mode and complete ODT provisioning checks
  ([d40bc2f](https://github.com/adnoctem/PSFoundation/commit/d40bc2fc064a1e8b4092e349f38fffc99d2bb2f0))
- **src:** support Windows 10 22H2 Office deployments
  ([525e09d](https://github.com/adnoctem/PSFoundation/commit/525e09d32e935f31e1852bd3fca140fa671ba6f6))

## [1.8.0](https://github.com/adnoctem/PSFoundation/compare/v1.7.5...v1.8.0) (2026-09-29)

### Features

- **src:** derive Click-to-Run languages from native registration
  ([96a3892](https://github.com/adnoctem/PSFoundation/commit/96a389207c6d29efdb62a61ab3a5803af936343a))

## [1.7.5](https://github.com/adnoctem/PSFoundation/compare/v1.7.4...v1.7.5) (2026-09-29)

### Bug Fixes

- **src:** recognize the native Click-to-Run Licensing Component
  ([46386e9](https://github.com/adnoctem/PSFoundation/commit/46386e9dba0a9db471227f3431029f324dfaa9d7))

## [1.7.4](https://github.com/adnoctem/PSFoundation/compare/v1.7.3...v1.7.4) (2026-09-29)

### Bug Fixes

- **src:** classify Office patch registrations and MSI App Paths
  ([7858057](https://github.com/adnoctem/PSFoundation/commit/7858057ff12e6209e0ce5ba4b0b470c641a66c16))
- **tools:** honour SkipChecksums in the prepare phase
  ([ac480a0](https://github.com/adnoctem/PSFoundation/commit/ac480a0eaa9aef8f4cbfc9043108091f7c6a201c))
- **tools:** keep the release manifest rewrite format-clean
  ([450c06d](https://github.com/adnoctem/PSFoundation/commit/450c06d7b84810a741b850119fd3c58213fab5c0))

## [1.7.3](https://github.com/adnoctem/PSFoundation/compare/v1.7.2...v1.7.3) (2026-09-29)

### Bug Fixes

- clear formatting inconsistencies ([886dbe7](https://github.com/adnoctem/PSFoundation/commit/886dbe7044f40cad8e507190cb0070dd526d4b11))
- **src:** handle sparse Office inventory under strict mode
  ([3493f7b](https://github.com/adnoctem/PSFoundation/commit/3493f7bf339a88f44042e4d2c1c0d3b3caef4688))

## [1.7.2](https://github.com/adnoctem/PSFoundation/compare/v1.7.1...v1.7.2) (2026-09-28)

### Bug Fixes

- **src:** recognize current Microsoft ODT bootstrapper metadata
  ([9b3769a](https://github.com/adnoctem/PSFoundation/commit/9b3769a8020b7690ed82a84cdd8b17ee978ac566))

## [1.7.1](https://github.com/adnoctem/PSFoundation/compare/v1.7.0...v1.7.1) (2026-09-28)

### Bug Fixes

- **src:** correct Outlook repair discovery and normalize filesystem paths
  ([3c5af7e](https://github.com/adnoctem/PSFoundation/commit/3c5af7e98c5c3a7a963b4e10ac0b48c01c5a1ceb))

## [1.7.0](https://github.com/adnoctem/PSFoundation/compare/v1.6.2...v1.7.0) (2026-09-28)

### Features

- **src:** add identity-based Outlook folder planning
  ([d11c7b1](https://github.com/adnoctem/PSFoundation/commit/d11c7b13d1294089d91f56042cf949d4c0aeddc7))

## [1.6.2](https://github.com/adnoctem/PSFoundation/compare/v1.6.1...v1.6.2) (2026-09-28)

### Bug Fixes

- **src:** handle empty and single-item collections under strict mode
  ([f14815b](https://github.com/adnoctem/PSFoundation/commit/f14815bb5f6e4759724bf1d2039fd9d8c43ad4c8))

## [1.6.1](https://github.com/adnoctem/PSFoundation/compare/v1.6.0...v1.6.1) (2026-09-28)

### Bug Fixes

- **src:** enable scoped Office 2007 migration pilots
  ([d47639a](https://github.com/adnoctem/PSFoundation/commit/d47639a79c1b9adcc8a9bceaf1dc7e562cadbe71))

## [1.6.0](https://github.com/adnoctem/PSFoundation/compare/v1.5.0...v1.6.0) (2026-09-27)

### Features

- **src:** add scoped Office deployment APIs
  ([3efcf90](https://github.com/adnoctem/PSFoundation/commit/3efcf90968ba6c67cc1dbdae03e2e7d4c1aa53c2))

### Bug Fixes

- **src:** correct Office inventory and block unverifiable deployments
  ([a860691](https://github.com/adnoctem/PSFoundation/commit/a8606916f8306bd35eb9aad901a3911850cbd0e0))

## [1.5.0](https://github.com/adnoctem/PSFoundation/compare/v1.4.0...v1.5.0) (2026-09-17)

### Features

- **src:** add registry restoration and operation diagnostics
  ([d289bda](https://github.com/adnoctem/PSFoundation/commit/d289bdadbd6a64b96bb109c3949cd75d0ca86769))

## [1.4.0](https://github.com/adnoctem/PSFoundation/compare/v1.3.0...v1.4.0) (2026-09-16)

### Features

- **src:** add policy conversion and deployment error guidance
  ([2961477](https://github.com/adnoctem/PSFoundation/commit/2961477846997d1deabe9299dafc9c1163e99267))

### Bug Fixes

- **tools:** scope dependency bumps to declared version fields
  ([8507ee9](https://github.com/adnoctem/PSFoundation/commit/8507ee9f027016fe46a53eced61f949a7e83bb89))

## [1.3.0](https://github.com/adnoctem/PSFoundation/compare/v1.2.0...v1.3.0) (2026-08-21)

### Features

- **src:** add Get-TransportMessageId for transport header parsing
  ([95e40ef](https://github.com/adnoctem/PSFoundation/commit/95e40ef4a270978dd79b3aa624b4e3fc19dda640))

## [1.2.0](https://github.com/adnoctem/PSFoundation/compare/v1.1.0...v1.2.0) (2026-08-18)

### Features

- **src:** add font, service, and scheduled-task state primitives
  ([8c8cce7](https://github.com/adnoctem/PSFoundation/commit/8c8cce7f2942c30780be01274f2e678d8bdbf44e))

## [1.1.0](https://github.com/adnoctem/PSFoundation/compare/v1.0.1...v1.1.0) (2026-08-18)

### Features

- **src:** add ownership, provisioning, and remote diagnostics primitives
  ([2e240bd](https://github.com/adnoctem/PSFoundation/commit/2e240bdd2fabb916905b96e2bfb91383b4077ad7))

### Bug Fixes

- **src:** make registry queries StrictMode-safe and harden test isolation
  ([f55b0c8](https://github.com/adnoctem/PSFoundation/commit/f55b0c8d0d4bdaa5fe269571afc7537cd7d67df4))
- **src:** tolerate missing optional OS registry values under StrictMode
  ([786fd91](https://github.com/adnoctem/PSFoundation/commit/786fd91f0cb6bf9fe24357972383a46ff636e0bc))

## [1.0.1](https://github.com/adnoctem/PSFoundation/compare/v1.0.0...v1.0.1) (2026-08-11)

### Bug Fixes

- **psd1:** correct invalid dependencies preventing installation
  ([e87e3a5](https://github.com/adnoctem/PSFoundation/commit/e87e3a5c33eeaabcbad6f1561baea952e7e88dfa))

## 1.0.0 (2026-08-10)

### Features

- **maintenance:** add module introspection, cleanup, and install functions
  ([f291b6b](https://github.com/adnoctem/PSFoundation/commit/f291b6bc47e3fe5f72f83da47e68052b5da533c5))
