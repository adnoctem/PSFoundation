# Registry policy fixtures

`ordered-dword.hex` is hand-authored from the
[MS-GPREG format specification](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-gpreg/5c092c22-bf6b-4e7f-b180-b20743d368f5),
independently of the production encoder. Whitespace separates hexadecimal bytes. It contains a PReg version 1 header and two DWORD
instructions for `Software\Test`, value `Flag`, first `4294967295`, then `1`. Both records must survive in that order.

This synthetic fixture contains no machine configuration. Tests materialize it in Pester's TestDrive. It is not the consumer's private
27-entry policy file, which is not available in this repository.
