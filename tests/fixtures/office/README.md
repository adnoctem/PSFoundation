# Office registry observations

These reduced fixtures reproduce registry shapes observed on Office Enterprise 2007 and Office Standard 2019 installations. They retain
public product/component names, product codes, versions, language identifiers, and registry views. The Click-to-Run active-configuration
identifier is synthetic. Machine identity, host configuration, timestamps, user data, activation data, and original report files are
excluded.

The fixtures omit missing keys, patch registrations, and non-GUID uninstall wrappers. The collector did not capture uninstall commands;
wrapper handling is covered separately with synthetic records in `office.Tests.ps1`. These are classification examples, not complete native
inventories or proof of installed-language completeness. No installer was run to produce these fixtures.

The 2019 fixture intentionally has no installed-build inventory key. The active product's `de-de` and `x-none` leaves and the corresponding
shared culture leaves retain observed `Version=16.0.10417.20208` values from the x86 report. Its active-configuration GUID is synthetic; the
channel URL is reduced to its channel identifier. The product-resource fallback requires agreeing versions for every listed product
resource; shared culture leaves and configuration telemetry cannot establish the product build by themselves. Tests also model the x64
sample's agreeing `16.0.10417.20211` versions and deliberately inconsistent registrations. The fixture does not establish application health
or full resource completeness.

The imported-module closeout cases reuse these reduced shapes and explicitly construct synthetic changes: empty/missing/malformed exclusion
values, English/German primary-language orders, multiple Click-to-Run products, unknown fields, retained Visio/Project products and a
different source product ID. Those mutations exercise contracts; they are not captures of those workstation combinations. Machine identity,
ACL SIDs, media paths, and recovery journals used by tracked tests are synthetic.

The accepted Standard 2019 x86-to-x64 continuation supplied the observed target build/language and no-op result shape. Tests recreate the
relevant configuration and inventory without copying the private journal or identity. The different-source-ID repeat is synthetic.
Historical schema-1 recovery tests cover JSON round trips, unchanged recorded fingerprints, pre-launch/completed-removal checkpoints,
complete targets and explicitly unsupported uncertain states through a full module import. Native recovery acceptance is limited to the
separately recorded post-removal continuation; a mock checkpoint is not an interruption experiment.
