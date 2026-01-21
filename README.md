# Rig Profiles

Hardware profiling script for USB compatibility testing. Collects system information and submits it to help identify compatibility patterns with MH4 boards.

## Quick Start

Run in PowerShell:

```powershell
irm https://tools.mariusheier.com/rig.ps1 | iex
```

## What It Collects

- System/motherboard information
- CPU and memory configuration
- USB controller topology (CPU-direct vs chipset)
- Power management settings
- Driver versions

## Verification

All releases include SHA256 checksums. To verify:

```powershell
(Get-FileHash rig.ps1 -Algorithm SHA256).Hash
```

Compare with the checksum in the [release notes](https://github.com/MariusHeier/rig-profiles/releases).

## View Submissions

Browse community hardware profiles at [tools.mariusheier.com/rig.html](https://tools.mariusheier.com/rig.html)

## License

MIT
