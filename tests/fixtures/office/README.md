# Office registry observations

These reduced fixtures reproduce registry shapes observed on Office Enterprise 2007 and Office Standard 2019 installations. They retain
public product/component names, product codes, versions, language identifiers, and registry views. The Click-to-Run active-configuration
identifier is synthetic. Machine identity, host configuration, timestamps, user data, activation data, and original report files are
excluded.

The fixtures omit missing keys, patch registrations, and non-GUID uninstall wrappers. The collector did not capture uninstall commands;
wrapper handling is covered separately with synthetic records in `office.Tests.ps1`. These are classification examples, not complete native
inventories or proof of installed-language completeness. No installer was run to produce these fixtures.

The 2019 fixture intentionally has no installed-build inventory key. Its configuration version must not become verified installation
evidence. Active product language registrations remain candidates; the initial shell language and full resource completeness remain unknown.
