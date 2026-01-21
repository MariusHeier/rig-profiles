#requires -Version 5.1
<#
.SYNOPSIS
    USB Rig Checker v3.0 - Hardware Profiling & Diagnostics
.DESCRIPTION
    Collects hardware data for USB compatibility testing.
    Captures: identity, chipset, USB controllers, CPU-direct detection, power settings.
    Uses comprehensive USB controller database from Linux kernel xhci-pci.c + pci.ids.
.PARAMETER HubHops
    Run in hub hops mode only - just show HID devices and their hub hop count.
#>

[CmdletBinding()]
param(
    [string]$OutputFileName = "",
    [switch]$Quiet,
    [switch]$HubHops
)

$ErrorActionPreference = 'SilentlyContinue'

# ==========================================
# COMPREHENSIVE USB CONTROLLER DATABASE
# (from Linux kernel xhci-pci.c + pci.ids)
# ==========================================

# Intel CPU-Integrated (CHIP 0 - LOWEST LATENCY)
$IntelCpuIntegrated = @{
    '8a13' = @{ Name = "Ice Lake Thunderbolt 3 USB Controller"; Platform = "Ice Lake (10th Gen)"; USB = "USB 3.2/TB3"; Year = "2019" }
    '9a13' = @{ Name = "Tiger Lake-LP Thunderbolt 4 USB Controller"; Platform = "Tiger Lake (11th Gen)"; USB = "USB4/TB4"; Year = "2020" }
    '9a17' = @{ Name = "Tiger Lake-H Thunderbolt 4 USB Controller"; Platform = "Tiger Lake-H (11th Gen)"; USB = "USB4/TB4"; Year = "2021" }
    '461e' = @{ Name = "Alder Lake-P Thunderbolt 4 USB Controller"; Platform = "Alder Lake (12th Gen)"; USB = "USB4/TB4"; Year = "2022" }
    '464e' = @{ Name = "Alder Lake-N Processor USB 3.2 xHCI Controller"; Platform = "Alder Lake-N"; USB = "USB 3.2"; Year = "2023" }
    'a71e' = @{ Name = "Raptor Lake-P Thunderbolt 4 USB Controller"; Platform = "Raptor Lake (13th Gen)"; USB = "USB4/TB4"; Year = "2023" }
    '7ec0' = @{ Name = "Meteor Lake-P Thunderbolt 4 USB Controller"; Platform = "Meteor Lake (Core Ultra)"; USB = "USB4/TB4"; Year = "2024" }
    'a831' = @{ Name = "Lunar Lake-M Thunderbolt 4 USB Controller"; Platform = "Lunar Lake"; USB = "USB4/TB4"; Year = "2024" }
}

# Intel PCH/Chipset (CHIP 1)
$IntelPch = @{
    '7f6e' = @{ Name = "800 Series PCH USB 3.1 xHCI HC"; Platform = "800 Series PCH"; USB = "USB 3.1"; Year = "2024" }
    '7a60' = @{ Name = "Raptor Lake USB 3.2 Gen 2x2 XHCI Host Controller"; Platform = "700 Series PCH"; USB = "USB 3.2 Gen 2x2 (20Gbps)"; Year = "2023" }
    '7a61' = @{ Name = "Raptor Lake USB 3.2 Gen 1x1 xDCI Device Controller"; Platform = "700 Series PCH"; USB = "USB 3.2 Gen 1"; Year = "2023" }
    '7ae0' = @{ Name = "Alder Lake-S PCH USB 3.2 Gen 2x2 XHCI Controller"; Platform = "600 Series PCH (Desktop)"; USB = "USB 3.2 Gen 2x2 (20Gbps)"; Year = "2021" }
    '51ed' = @{ Name = "Alder Lake PCH USB 3.2 xHCI Host Controller"; Platform = "600 Series PCH"; USB = "USB 3.2"; Year = "2022" }
    '54ed' = @{ Name = "Alder Lake-N PCH USB 3.2 Gen 2x1 xHCI Host Controller"; Platform = "Alder Lake-N PCH"; USB = "USB 3.2 Gen 2 (10Gbps)"; Year = "2023" }
    '7e7d' = @{ Name = "Meteor Lake-P USB 3.2 Gen 2x1 xHCI Host Controller"; Platform = "Meteor Lake PCH"; USB = "USB 3.2 Gen 2"; Year = "2024" }
    '777d' = @{ Name = "Arrow Lake USB 3.2 xHCI Controller"; Platform = "Arrow Lake"; USB = "USB 3.2"; Year = "2024" }
    'a87d' = @{ Name = "Lunar Lake-M USB 3.2 Gen 2x1 xHCI Host Controller"; Platform = "Lunar Lake PCH"; USB = "USB 3.2 Gen 2"; Year = "2024" }
    'a0ed' = @{ Name = "Tiger Lake-LP USB 3.2 Gen 2x1 xHCI Host Controller"; Platform = "500 Series PCH"; USB = "USB 3.2 Gen 2 (10Gbps)"; Year = "2020" }
    '43ed' = @{ Name = "Tiger Lake-H USB 3.2 Gen 2x1 xHCI Host Controller"; Platform = "500 Series PCH-H"; USB = "USB 3.2 Gen 2"; Year = "2021" }
    'a3af' = @{ Name = "Comet Lake PCH-V USB Controller"; Platform = "400 Series PCH"; USB = "USB 3.1"; Year = "2020" }
    '02ed' = @{ Name = "Comet Lake PCH-LP USB 3.1 xHCI Host Controller"; Platform = "400 Series PCH-LP"; USB = "USB 3.1"; Year = "2020" }
    '06ed' = @{ Name = "Comet Lake USB 3.1 xHCI Host Controller"; Platform = "400 Series PCH"; USB = "USB 3.1"; Year = "2020" }
    'a36d' = @{ Name = "Cannon Lake PCH USB 3.1 xHCI Host Controller"; Platform = "300 Series PCH"; USB = "USB 3.1"; Year = "2018" }
    '9ded' = @{ Name = "Cannon Point-LP USB 3.1 xHCI Controller"; Platform = "300 Series PCH-LP"; USB = "USB 3.1"; Year = "2018" }
    'a2af' = @{ Name = "200 Series/Z370 Chipset Family USB 3.0 xHCI Controller"; Platform = "200 Series PCH"; USB = "USB 3.0"; Year = "2017" }
    'a12f' = @{ Name = "100 Series/C230 Series Chipset Family USB 3.0 xHCI Controller"; Platform = "100 Series PCH"; USB = "USB 3.0"; Year = "2015" }
    '9d2f' = @{ Name = "Sunrise Point-LP USB 3.0 xHCI Controller"; Platform = "100 Series PCH-LP"; USB = "USB 3.0"; Year = "2015" }
    '8cb1' = @{ Name = "9 Series Chipset Family USB xHCI Controller"; Platform = "9 Series PCH"; USB = "USB 3.0"; Year = "2014" }
    '9cb1' = @{ Name = "Wildcat Point-LP USB xHCI Controller"; Platform = "9 Series PCH-LP"; USB = "USB 3.0"; Year = "2014" }
    '8c31' = @{ Name = "8 Series/C220 Series Chipset Family USB xHCI"; Platform = "8 Series PCH"; USB = "USB 3.0"; Year = "2013" }
    '9c31' = @{ Name = "8 Series USB xHCI HC"; Platform = "8 Series PCH-LP"; USB = "USB 3.0"; Year = "2013" }
    '1e31' = @{ Name = "7 Series/C210 Series Chipset Family USB xHCI Host Controller"; Platform = "7 Series PCH"; USB = "USB 3.0"; Year = "2012" }
    '8d31' = @{ Name = "C610/X99 series chipset USB xHCI Host Controller"; Platform = "X99/C610 (HEDT/Server)"; USB = "USB 3.0"; Year = "2014" }
    'a1af' = @{ Name = "C620 Series Chipset Family USB 3.0 xHCI Controller"; Platform = "C620 (Server)"; USB = "USB 3.0"; Year = "2017" }
}

# Intel Thunderbolt (CHIP 0 - CPU-attached)
$IntelThunderbolt = @{
    '5782' = @{ Name = "JHL9580 Thunderbolt 5 USB Controller"; Platform = "Barlow Ridge Host 80G"; USB = "USB4/TB5 (80Gbps)"; Year = "2024" }
    '5785' = @{ Name = "JHL9540 Thunderbolt 4 USB Controller"; Platform = "Barlow Ridge Host 40G"; USB = "USB4/TB4 (40Gbps)"; Year = "2024" }
    '5787' = @{ Name = "JHL9480 Thunderbolt 5 USB Controller"; Platform = "Barlow Ridge Hub 80G"; USB = "USB4/TB5 (80Gbps)"; Year = "2024" }
    '57a5' = @{ Name = "JHL9440 Thunderbolt 4 USB Controller"; Platform = "Barlow Ridge Hub 40G"; USB = "USB4/TB4 (40Gbps)"; Year = "2024" }
    '1138' = @{ Name = "Thunderbolt 4 USB Controller [Maple Ridge 4C]"; Platform = "Maple Ridge 4C"; USB = "USB4/TB4 (40Gbps)"; Year = "2020" }
    '1135' = @{ Name = "Thunderbolt 4 USB Controller [Maple Ridge 2C]"; Platform = "Maple Ridge 2C"; USB = "USB4/TB4 (40Gbps)"; Year = "2020" }
    '0b27' = @{ Name = "Thunderbolt 4 USB Controller [Goshen Ridge]"; Platform = "Goshen Ridge"; USB = "USB4/TB4 (40Gbps)"; Year = "2020" }
    '15e9' = @{ Name = "JHL7540 Thunderbolt 3 USB Controller [Titan Ridge 2C]"; Platform = "Titan Ridge 2C"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2018" }
    '15ec' = @{ Name = "JHL7540 Thunderbolt 3 USB Controller [Titan Ridge 4C]"; Platform = "Titan Ridge 4C"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2018" }
    '15f0' = @{ Name = "JHL7440 Thunderbolt 3 USB Controller [Titan Ridge DD]"; Platform = "Titan Ridge DD"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2018" }
    '15b5' = @{ Name = "DSL6340 USB 3.1 Controller [Alpine Ridge 2C]"; Platform = "Alpine Ridge 2C"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2015" }
    '15b6' = @{ Name = "DSL6540 USB 3.1 Controller [Alpine Ridge 4C]"; Platform = "Alpine Ridge 4C"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2015" }
    '15c1' = @{ Name = "JHL6240 Thunderbolt 3 USB 3.1 Controller [Alpine Ridge LP]"; Platform = "Alpine Ridge LP"; USB = "USB 3.1/TB3"; Year = "2016" }
    '15d4' = @{ Name = "JHL6540 Thunderbolt 3 USB Controller [Alpine Ridge 4C]"; Platform = "Alpine Ridge 4C C-step"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2016" }
    '15db' = @{ Name = "JHL6340 Thunderbolt 3 USB 3.1 Controller [Alpine Ridge 2C]"; Platform = "Alpine Ridge 2C C-step"; USB = "USB 3.1/TB3 (40Gbps)"; Year = "2016" }
}

# AMD CPU-Integrated (CHIP 0 - LOWEST LATENCY)
$AmdCpuIntegrated = @{
    # Raphael / Granite Ridge (Ryzen 7000/9000)
    '15b6' = @{ Name = "Raphael/Granite Ridge USB 3.1 xHCI"; Platform = "Ryzen 7000/9000 Desktop (AM5)"; USB = "USB 3.1"; Year = "2022" }
    '15b7' = @{ Name = "Raphael/Granite Ridge USB 3.1 xHCI"; Platform = "Ryzen 7000/9000 Desktop (AM5)"; USB = "USB 3.1"; Year = "2022" }
    '15b8' = @{ Name = "Raphael/Granite Ridge USB 2.0 xHCI"; Platform = "Ryzen 7000/9000 Desktop (AM5)"; USB = "USB 2.0"; Year = "2022" }
    # Strix Halo (Zen 5)
    '1587' = @{ Name = "Strix Halo USB 3.1 xHCI"; Platform = "Strix Halo (Zen 5)"; USB = "USB 3.1"; Year = "2024" }
    '1588' = @{ Name = "Strix Halo USB 3.1 xHCI"; Platform = "Strix Halo (Zen 5)"; USB = "USB 3.1"; Year = "2024" }
    '1589' = @{ Name = "Strix Halo USB 3.1 xHCI"; Platform = "Strix Halo (Zen 5)"; USB = "USB 3.1"; Year = "2024" }
    '158b' = @{ Name = "Strix Halo USB 3.1 xHCI"; Platform = "Strix Halo (Zen 5)"; USB = "USB 3.1"; Year = "2024" }
    '158d' = @{ Name = "Strix Halo USB4 Host Router"; Platform = "Strix Halo (Zen 5)"; USB = "USB4"; Year = "2024" }
    '158e' = @{ Name = "Strix Halo USB4 Host Router"; Platform = "Strix Halo (Zen 5)"; USB = "USB4"; Year = "2024" }
    # Rembrandt (Ryzen 6000 Mobile)
    '161a' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '161b' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '161c' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '161d' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '161e' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '161f' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '15d6' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '15d7' = @{ Name = "Rembrandt USB4 XHCI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4"; Year = "2022" }
    '162e' = @{ Name = "Rembrandt USB4/Thunderbolt NHI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4/TB"; Year = "2022" }
    '162f' = @{ Name = "Rembrandt USB4/Thunderbolt NHI controller"; Platform = "Ryzen 6000 Mobile (Zen 3+)"; USB = "USB4/TB"; Year = "2022" }
    # Phoenix (Ryzen 7040)
    '15c4' = @{ Name = "Phoenix USB4/Thunderbolt NHI controller"; Platform = "Ryzen 7040 Mobile (Zen 4)"; USB = "USB4/TB"; Year = "2023" }
    '15c5' = @{ Name = "Phoenix USB4/Thunderbolt NHI controller"; Platform = "Ryzen 7040 Mobile (Zen 4)"; USB = "USB4/TB"; Year = "2023" }
    '1668' = @{ Name = "Pink Sardine USB4/Thunderbolt NHI controller"; Platform = "Pink Sardine"; USB = "USB4/TB"; Year = "2023" }
    '1669' = @{ Name = "Pink Sardine USB4/Thunderbolt NHI controller"; Platform = "Pink Sardine"; USB = "USB4/TB"; Year = "2023" }
    # Renoir / Cezanne (Ryzen 4000/5000 APU)
    '1639' = @{ Name = "Renoir/Cezanne USB 3.1"; Platform = "Ryzen 4000/5000 APU (Zen 2/3)"; USB = "USB 3.1"; Year = "2020" }
    # Raven Ridge / Picasso (Ryzen 2000/3000 APU)
    '15e0' = @{ Name = "Raven USB 3.1"; Platform = "Ryzen 2000 APU (Zen)"; USB = "USB 3.1"; Year = "2018" }
    '15e1' = @{ Name = "Raven USB 3.1"; Platform = "Ryzen 2000 APU (Zen)"; USB = "USB 3.1"; Year = "2018" }
    '15e5' = @{ Name = "Raven2 USB 3.1"; Platform = "Ryzen 3000 APU (Zen+)"; USB = "USB 3.1"; Year = "2019" }
    # Matisse / Vermeer (Ryzen 3000/5000 Desktop)
    '149c' = @{ Name = "Matisse USB 3.0 Host Controller"; Platform = "Ryzen 3000/5000 Desktop (Zen 2/3)"; USB = "USB 3.0"; Year = "2019" }
    '148c' = @{ Name = "Starship USB 3.0 Host Controller"; Platform = "EPYC Rome / Threadripper 3rd Gen"; USB = "USB 3.0"; Year = "2019" }
    # Zeppelin (Ryzen 1000)
    '145f' = @{ Name = "Zeppelin USB 3.0 xHCI Compliant Host Controller"; Platform = "Ryzen 1000 (Zen)"; USB = "USB 3.0"; Year = "2017" }
    '145c' = @{ Name = "Family 17h USB 3.0 Host Controller"; Platform = "Ryzen 1000 (Zen)"; USB = "USB 3.0"; Year = "2017" }
    # Van Gogh (Steam Deck)
    '162c' = @{ Name = "VanGogh USB2"; Platform = "Steam Deck (Van Gogh)"; USB = "USB 2.0"; Year = "2022" }
    '163a' = @{ Name = "VanGogh USB0"; Platform = "Steam Deck (Van Gogh)"; USB = "USB 3.1"; Year = "2022" }
    '163b' = @{ Name = "VanGogh USB1"; Platform = "Steam Deck (Van Gogh)"; USB = "USB 3.1"; Year = "2022" }
    # Other
    '15d4' = @{ Name = "FireFlight USB 3.1"; Platform = "FireFlight"; USB = "USB 3.1"; Year = "2020" }
    '15d5' = @{ Name = "FireFlight USB 3.1"; Platform = "FireFlight"; USB = "USB 3.1"; Year = "2020" }
    '13ed' = @{ Name = "Ariel USB 3.1 Type C (Gen2 + DP Alt)"; Platform = "Ariel"; USB = "USB 3.1 Gen 2"; Year = "2020" }
    '13ee' = @{ Name = "Ariel USB 3.1 Type A (Gen2 x 2 ports)"; Platform = "Ariel"; USB = "USB 3.1 Gen 2"; Year = "2020" }
    '1557' = @{ Name = "Turin USB 3.1 xHCI"; Platform = "EPYC Turin"; USB = "USB 3.1"; Year = "2024" }
}

# AMD Chipset (CHIP 1)
$AmdChipset = @{
    '43fc' = @{ Name = "800 Series Chipset USB 3.x XHCI Controller"; Platform = "X870/B850 (AM5)"; USB = "USB 3.2"; Year = "2024" }
    '43fd' = @{ Name = "800 Series Chipset USB 3.x XHCI Controller"; Platform = "X870/B850 (AM5)"; USB = "USB 3.2"; Year = "2024" }
    '43f7' = @{ Name = "600 Series Chipset USB 3.2 Controller"; Platform = "X670/B650 (AM5)"; USB = "USB 3.2"; Year = "2022" }
    '43ee' = @{ Name = "500 Series Chipset USB 3.1 XHCI Controller"; Platform = "X570/B550 (AM4)"; USB = "USB 3.1"; Year = "2019" }
    '43ec' = @{ Name = "A520 Series Chipset USB 3.1 XHCI Controller"; Platform = "A520 (AM4)"; USB = "USB 3.1"; Year = "2020" }
    '43d5' = @{ Name = "400 Series Chipset USB 3.1 xHCI Compliant Host Controller"; Platform = "X470/B450 (AM4)"; USB = "USB 3.1"; Year = "2018" }
    '43b9' = @{ Name = "X370 Series Chipset USB 3.1 xHCI Controller"; Platform = "X370 (AM4)"; USB = "USB 3.1"; Year = "2017" }
    '43ba' = @{ Name = "X399 Series Chipset USB 3.1 xHCI Controller"; Platform = "X399 (Threadripper)"; USB = "USB 3.1"; Year = "2017" }
    '43bb' = @{ Name = "300 Series Chipset USB 3.1 xHCI Controller"; Platform = "B350 (AM4)"; USB = "USB 3.1"; Year = "2017" }
    '43bc' = @{ Name = "A320 USB 3.1 XHCI Host Controller"; Platform = "A320 (AM4)"; USB = "USB 3.1"; Year = "2017" }
    '7814' = @{ Name = "FCH USB XHCI Controller"; Platform = "Legacy FCH"; USB = "USB 3.0"; Year = "2013" }
    '7812' = @{ Name = "FCH USB XHCI Controller"; Platform = "Legacy FCH"; USB = "USB 3.0"; Year = "2012" }
}

# Third-party add-in cards (CHIP 1+)
$ThirdPartyControllers = @{
    # ASMedia (VID: 1b21)
    '1b21:1042' = @{ Name = "ASM1042 SuperSpeed USB Host Controller"; Vendor = "ASMedia"; USB = "USB 3.0 (5Gbps)"; Year = "2011" }
    '1b21:1142' = @{ Name = "ASM1042A USB 3.0 Host Controller"; Vendor = "ASMedia"; USB = "USB 3.0 (5Gbps)"; Year = "2013" }
    '1b21:1242' = @{ Name = "ASM1142 USB 3.1 Host Controller"; Vendor = "ASMedia"; USB = "USB 3.1 Gen 2 (10Gbps)"; Year = "2015" }
    '1b21:1343' = @{ Name = "ASM1143 USB 3.1 Host Controller"; Vendor = "ASMedia"; USB = "USB 3.1 Gen 2 (10Gbps)"; Year = "2017" }
    '1b21:2142' = @{ Name = "ASM2142/ASM3142 USB 3.1 Host Controller"; Vendor = "ASMedia"; USB = "USB 3.1 Gen 2 (10Gbps)"; Year = "2016" }
    '1b21:3042' = @{ Name = "ASM3042 USB 3.2 Gen 1 xHCI Controller"; Vendor = "ASMedia"; USB = "USB 3.2 Gen 1 (5Gbps)"; Year = "2019" }
    '1b21:3142' = @{ Name = "ASM3142 USB 3.2 Gen 2x1 xHCI Controller"; Vendor = "ASMedia"; USB = "USB 3.2 Gen 2 (10Gbps)"; Year = "2019" }
    '1b21:3242' = @{ Name = "ASM3242 USB 3.2 Host Controller"; Vendor = "ASMedia"; USB = "USB 3.2 Gen 2x2 (20Gbps)"; Year = "2020" }
    '1b21:2425' = @{ Name = "ASM4242 USB 4 / Thunderbolt 3 Host Router"; Vendor = "ASMedia"; USB = "USB4/TB3 (40Gbps)"; Year = "2022" }
    '1b21:2426' = @{ Name = "ASM4242 USB 3.2 xHCI Controller"; Vendor = "ASMedia"; USB = "USB 3.2"; Year = "2022" }
    # VIA (VID: 1106)
    '1106:3483' = @{ Name = "VL805/806 xHCI USB 3.0 Controller"; Vendor = "VIA"; USB = "USB 3.0 (5Gbps)"; Year = "2014" }
    '1106:3432' = @{ Name = "VL800/801 xHCI USB 3.0 Controller"; Vendor = "VIA"; USB = "USB 3.0 (5Gbps)"; Year = "2012" }
    # Fresco Logic (VID: 1b73)
    '1b73:1000' = @{ Name = "FL1000G USB 3.0 Host Controller"; Vendor = "Fresco Logic"; USB = "USB 3.0 (5Gbps)"; Year = "2010" }
    '1b73:1009' = @{ Name = "FL1009 USB 3.0 Host Controller"; Vendor = "Fresco Logic"; USB = "USB 3.0 (5Gbps)"; Year = "2011" }
    '1b73:1100' = @{ Name = "FL1100 USB 3.0 Host Controller"; Vendor = "Fresco Logic"; USB = "USB 3.0 (5Gbps)"; Year = "2012" }
    '1b73:1400' = @{ Name = "FL1400 USB 3.0 Host Controller"; Vendor = "Fresco Logic"; USB = "USB 3.0 (5Gbps)"; Year = "2014" }
    # Etron (VID: 1b6f)
    '1b6f:7023' = @{ Name = "EJ168 USB 3.0 Host Controller"; Vendor = "Etron"; USB = "USB 3.0 (5Gbps)"; Year = "2011" }
    '1b6f:7052' = @{ Name = "EJ188/EJ198 USB 3.0 Host Controller"; Vendor = "Etron"; USB = "USB 3.0 (5Gbps)"; Year = "2013" }
    # Renesas (VID: 1912)
    '1912:0014' = @{ Name = "uPD720201 USB 3.0 Host Controller"; Vendor = "Renesas"; USB = "USB 3.0 (5Gbps)"; Year = "2011" }
    '1912:0015' = @{ Name = "uPD720202 USB 3.0 Host Controller"; Vendor = "Renesas"; USB = "USB 3.0 (5Gbps)"; Year = "2012" }
}

# ==========================================
# HELPER: Get Controller Info from Database
# ==========================================
function Get-ControllerInfo {
    param([string]$Vid, [string]$Did)

    $vid = $Vid.ToLower()
    $did = $Did.ToLower()
    $key = "$vid`:$did"

    # Intel CPU-integrated = CHIP 0
    if ($vid -eq '8086' -and $IntelCpuIntegrated.ContainsKey($did)) {
        $data = $IntelCpuIntegrated[$did]
        return @{ Type = "CPU"; ChipCount = 0; Name = $data.Name; Platform = $data.Platform; USB = $data.USB }
    }

    # Intel Thunderbolt = CHIP 0 (CPU-attached)
    if ($vid -eq '8086' -and $IntelThunderbolt.ContainsKey($did)) {
        $data = $IntelThunderbolt[$did]
        return @{ Type = "TB"; ChipCount = 0; Name = $data.Name; Platform = $data.Platform; USB = $data.USB }
    }

    # Intel PCH = CHIP 1
    if ($vid -eq '8086' -and $IntelPch.ContainsKey($did)) {
        $data = $IntelPch[$did]
        return @{ Type = "Chipset"; ChipCount = 1; Name = $data.Name; Platform = $data.Platform; USB = $data.USB }
    }

    # AMD CPU-integrated = CHIP 0
    if ($vid -eq '1022' -and $AmdCpuIntegrated.ContainsKey($did)) {
        $data = $AmdCpuIntegrated[$did]
        return @{ Type = "CPU"; ChipCount = 0; Name = $data.Name; Platform = $data.Platform; USB = $data.USB }
    }

    # AMD Chipset = CHIP 1
    if ($vid -eq '1022' -and $AmdChipset.ContainsKey($did)) {
        $data = $AmdChipset[$did]
        return @{ Type = "Chipset"; ChipCount = 1; Name = $data.Name; Platform = $data.Platform; USB = $data.USB }
    }

    # Third-party
    if ($ThirdPartyControllers.ContainsKey($key)) {
        $data = $ThirdPartyControllers[$key]
        return @{ Type = "Addon"; ChipCount = 1; Name = $data.Name; Platform = "PCIe Add-in"; USB = $data.USB }
    }

    # Unknown third-party vendors
    $vendors = @{ '1b21' = 'ASMedia'; '1106' = 'VIA'; '1b73' = 'Fresco Logic'; '1912' = 'Renesas'; '1b6f' = 'Etron'; '104c' = 'Texas Instruments' }
    if ($vendors.ContainsKey($vid)) {
        return @{ Type = "Addon"; ChipCount = 1; Name = "$($vendors[$vid]) Controller"; Platform = "PCIe Add-in"; USB = "USB 3.x" }
    }

    # Intel unknown = assume PCH
    if ($vid -eq '8086') {
        return @{ Type = "Chipset"; ChipCount = 1; Name = "Intel USB Controller"; Platform = "Unknown PCH (DID:$did)"; USB = "USB 3.x" }
    }

    # AMD unknown = assume chipset
    if ($vid -eq '1022') {
        return @{ Type = "Chipset"; ChipCount = 1; Name = "AMD USB Controller"; Platform = "Unknown Chipset (DID:$did)"; USB = "USB 3.x" }
    }

    return @{ Type = "Unknown"; ChipCount = 1; Name = "Unknown Controller"; Platform = "VID:$vid DID:$did"; USB = "?" }
}

# ==========================================
# HELPER: Get Chipset from SMBus
# ==========================================
$ChipsetDB = @{
    "1022:790B" = "AMD_B550"
    "1022:790E" = "AMD_X570"
    "1022:43EB" = "AMD_B550"
    "1022:43EC" = "AMD_X570"
    "1022:43F7" = "AMD_X670"
    "1022:43FC" = "AMD_X870"
    "8086:A323" = "INTEL_Z390"
    "8086:A2A3" = "INTEL_Z270"
    "8086:7A60" = "INTEL_Z790"
    "8086:7AE0" = "INTEL_Z690"
}

function Get-Chipset {
    $smbus = Get-CimInstance Win32_PnPEntity | Where-Object { $_.Name -match "SMBus" } | Select-Object -First 1
    if ($smbus -and $smbus.PNPDeviceID -match "VEN_([0-9A-F]+)&DEV_([0-9A-F]+)") {
        $key = "$($matches[1]):$($matches[2])"
        if ($ChipsetDB[$key]) { return $ChipsetDB[$key] }
        return "UNKNOWN_$($matches[1])_$($matches[2])"
    }
    return "UNKNOWN"
}

function New-ShortId {
    $chars = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"  # No I, O (avoid confusion)
    $id = -join (1..4 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
    return "RIG-$id"
}

# ==========================================
# GRUVBOX COLORS
# ==========================================
$c = @{
    Header = "DarkCyan"
    Label  = "DarkYellow"
    Tip    = "Yellow"
    OK     = "DarkGreen"
    Error  = "DarkRed"
    Text   = "Gray"
    Value  = "White"
}

# ==========================================
# HUB HOPS STANDALONE MODE
# ==========================================
if ($HubHops) {
    Write-Host ""
    Write-Host "USB HUB HOPS CHECKER" -ForegroundColor $c.Header
    Write-Host ("=" * 60) -ForegroundColor $c.Header
    Write-Host ""

    function Get-UsbPathStandalone {
        param([string]$InstanceId)
        $path = [System.Collections.ArrayList]@()
        $currentId = $InstanceId
        $count = 0

        while ($currentId -and $count -lt 10) {
            $count++
            $dev = Get-PnpDevice -InstanceId $currentId -ErrorAction SilentlyContinue

            if ($currentId -match "ROOT_HUB") {
                if ($dev) { [void]$path.Add($dev.FriendlyName) }
                try {
                    $parentProp = Get-PnpDeviceProperty -InstanceId $currentId -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop
                    $ctrlDev = Get-PnpDevice -InstanceId $parentProp.Data -ErrorAction SilentlyContinue
                    if ($ctrlDev -and $ctrlDev.FriendlyName -match "Controller") {
                        [void]$path.Add($ctrlDev.FriendlyName)
                    }
                } catch { }
                break
            }

            if ($dev -and $dev.FriendlyName -match "Hub" -and $dev.FriendlyName -notmatch "Root") {
                [void]$path.Add($dev.FriendlyName)
            }

            try {
                $parentProp = Get-PnpDeviceProperty -InstanceId $currentId -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop
                $currentId = $parentProp.Data
            } catch { break }
        }
        return $path
    }

    $hidDevices = Get-PnpDevice -Class "Mouse", "Keyboard", "HIDClass" -Status OK -ErrorAction SilentlyContinue |
        Where-Object { $_.InstanceId -match "^USB\\" -or $_.InstanceId -match "^HID\\" }

    Write-Host "INPUT DEVICES (mice, keyboards, gamepads)" -ForegroundColor $c.Label
    Write-Host ""

    $found = $false
    foreach ($device in $hidDevices) {
        $instanceId = $device.InstanceId
        $usbParent = $instanceId
        if ($instanceId -match "^HID\\") {
            try {
                $parent = Get-PnpDeviceProperty -InstanceId $instanceId -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop
                $usbParent = $parent.Data
            } catch { }
        }

        # Skip generic HID sub-interfaces
        if ($device.FriendlyName -match "^HID-compliant|^USB Input Device$|^HID Keyboard Device$") { continue }

        $path = Get-UsbPathStandalone -InstanceId $usbParent
        $hops = $path.Count

        $devVid = "????"
        $devPid = "????"
        if ($usbParent -match "VID_([0-9A-F]{4})") { $devVid = $matches[1] }
        if ($usbParent -match "PID_([0-9A-F]{4})") { $devPid = $matches[1] }

        $hopColor = if ($hops -le 1) { $c.OK } elseif ($hops -ge 3) { $c.Tip } else { $c.Text }

        Write-Host "  $($device.FriendlyName)" -ForegroundColor $c.Value
        Write-Host "  VID:PID $devVid`:$devPid" -ForegroundColor $c.Text

        if ($path -and $path.Count -gt 0) {
            $pathItems = @($path)
            for ($i = 0; $i -lt $pathItems.Count; $i++) {
                $prefix = "    " + ("-" * $i) + ">"
                Write-Host "$prefix $($pathItems[$i])" -ForegroundColor $hopColor
            }
        } else {
            Write-Host "    -> (direct)" -ForegroundColor $c.OK
        }
        Write-Host ""
        $found = $true
    }

    if (-not $found) {
        Write-Host "  No HID devices found." -ForegroundColor $c.Text
    }

    Write-Host ("=" * 60) -ForegroundColor $c.Header
    Write-Host "Path shows: Device -> Hub(s) -> Root Hub -> Controller" -ForegroundColor $c.Text
    Write-Host "Fewer hubs = better (less latency, more stable power)" -ForegroundColor $c.Text
    Write-Host ""

    exit 0
}

# ==========================================
# QUESTIONNAIRE
# ==========================================
Write-Host ""
Write-Host "USB RIG CHECKER v3.0" -ForegroundColor $c.Header
Write-Host ("=" * 60) -ForegroundColor $c.Header
# Report type
Write-Host "What are you reporting?" -ForegroundColor $c.Label
Write-Host "  [1] Problem - something isn't working"
Write-Host "  [2] Working - confirming a working setup"
$reportTypeChoice = Read-Host "Select (1-2)"
$reportType = switch ($reportTypeChoice) {
    "1" { "problem" }
    "2" { "working" }
    default { "problem" }
}

# Board type
Write-Host ""
Write-Host "Board type:" -ForegroundColor $c.Label
Write-Host "  [1] MH4-Analog"
Write-Host "  [2] MH4-Digital"
Write-Host "  [3] Not sure"
$boardChoice = Read-Host "Select (1-3)"
$board = switch ($boardChoice) {
    "1" { "MH4-Analog" }
    "2" { "MH4-Digital" }
    "3" { "unknown" }
    default { "unspecified" }
}

# Cable type
Write-Host ""
Write-Host "Cable type (computer side):" -ForegroundColor $c.Label
Write-Host "  [1] USB-A"
Write-Host "  [2] USB-C"
Write-Host "  [3] Tested both"
$cableChoice = Read-Host "Select (1-3)"
$cable = switch ($cableChoice) {
    "1" { "usb-a" }
    "2" { "usb-c" }
    "3" { "both" }
    default { "unspecified" }
}
if ($reportType -eq "problem" -and $cableChoice -in @("1","2")) {
    $otherCable = if ($cableChoice -eq "1") { "USB-C" } else { "USB-A" }
    Write-Host "TIP: Also try a $otherCable cable if available." -ForegroundColor $c.Tip
}

# Port location
Write-Host ""
Write-Host "Port location:" -ForegroundColor $c.Label
Write-Host "  [1] Back"
Write-Host "  [2] Front"
Write-Host "  [3] Tested both"
$portChoice = Read-Host "Select (1-3)"
$port = switch ($portChoice) {
    "1" { "back" }
    "2" { "front" }
    "3" { "both" }
    default { "unspecified" }
}
if ($reportType -eq "problem" -and $portChoice -in @("1","2")) {
    $otherPort = if ($portChoice -eq "1") { "front" } else { "back" }
    Write-Host "TIP: Also try a $otherPort port to rule out port-specific issues." -ForegroundColor $c.Tip
}

# Poll rate - individual status for each
Write-Host ""
Write-Host "Poll rate testing (for each rate, enter: f=fail, o=ok, -=untested):" -ForegroundColor $c.Label

$pollRates = [ordered]@{
    "8000" = "untested"
    "4000" = "untested"
    "2000" = "untested"
    "1000" = "untested"
}

foreach ($rate in @("8000", "4000", "2000", "1000")) {
    $input = Read-Host "  ${rate}Hz [f/o/-]"
    $pollRates[$rate] = switch ($input.ToLower()) {
        "f" { "fail" }
        "o" { "ok" }
        "-" { "untested" }
        "" { "untested" }
        default { "untested" }
    }
}

# Tip if high rates fail but didn't test lower
$hasHighFail = $pollRates["8000"] -eq "fail" -or $pollRates["4000"] -eq "fail"
$hasLowUntested = $pollRates["1000"] -eq "untested"
if ($hasHighFail -and $hasLowUntested) {
    Write-Host "TIP: Try testing at 1000Hz to see if lower rates work." -ForegroundColor $c.Tip
}

# Software version - based on board type
Write-Host ""
Write-Host "MH4 software version:" -ForegroundColor $c.Label
Write-Host "  (Check https://update.mariusheier.com to find your version)" -ForegroundColor $c.Text

if ($board -eq "MH4-Digital") {
    Write-Host "  [1] v1.35_beta (latest)"
    Write-Host "  [2] v1.34_beta"
    Write-Host "  [3] v1.20_beta"
    Write-Host "  [4] v1.00 (initial release)"
    Write-Host "  [0] Don't know"
    $versionChoice = Read-Host "Select"
    $softwareVersion = switch ($versionChoice) {
        "1" { "v1.35_beta" }
        "2" { "v1.34_beta" }
        "3" { "v1.20_beta" }
        "4" { "v1.00" }
        default { "unknown" }
    }
} elseif ($board -eq "MH4-Analog") {
    Write-Host "  [1] v1.10_analog (latest)"
    Write-Host "  [2] v1.00-analog (initial release)"
    Write-Host "  [0] Don't know"
    $versionChoice = Read-Host "Select"
    $softwareVersion = switch ($versionChoice) {
        "1" { "v1.10_analog" }
        "2" { "v1.00-analog" }
        default { "unknown" }
    }
} else {
    Write-Host "  [0] Don't know / Not sure which board"
    Read-Host "Press Enter"
    $softwareVersion = "unknown"
}

# Description (problem or notes)
Write-Host ""
if ($reportType -eq "working") {
    Write-Host "Any notes about your working setup (enter to skip):" -ForegroundColor $c.Label
} else {
    Write-Host "Describe the problem:" -ForegroundColor $c.Label
}
$problemDescription = Read-Host

Write-Host ""
Write-Host "Collecting hardware profile..." -ForegroundColor $c.Text
Write-Host ""

# Email will be asked later, after showing known issues
$userEmail = $null

# ==========================================
# COLLECT HARDWARE DATA
# ==========================================

$cs = Get-CimInstance Win32_ComputerSystem
$se = Get-CimInstance Win32_SystemEnclosure
$bb = Get-CimInstance Win32_BaseBoard
$bios = Get-CimInstance Win32_BIOS
$cpu = Get-CimInstance Win32_Processor | Select-Object -First 1
$mem = Get-CimInstance Win32_PhysicalMemory
$memArray = Get-CimInstance Win32_PhysicalMemoryArray | Select-Object -First 1
$totalGB = [math]::Round($cs.TotalPhysicalMemory / 1GB, 1)
$usbCtrl = Get-CimInstance Win32_USBController
$disks = Get-CimInstance Win32_DiskDrive | Where-Object { $_.MediaType -match "Fixed|SSD|NVMe" }
$gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1
$allDevices = Get-PnpDevice -ErrorAction SilentlyContinue
$problemDevices = $allDevices | Where-Object { $_.Status -notin @('OK', 'Unknown') }
$os = Get-CimInstance Win32_OperatingSystem
$detectedChipset = Get-Chipset

# USB ports from SMBIOS
$usbPorts = @(Get-CimInstance Win32_PortConnector | Where-Object {
    $_.ExternalReferenceDesignator -match "USB" -or $_.InternalReferenceDesignator -match "USB"
} | ForEach-Object {
    if ($_.ExternalReferenceDesignator) { $_.ExternalReferenceDesignator } else { $_.InternalReferenceDesignator }
})

# Power settings
$powerPlan = powercfg /getactivescheme 2>&1 | Out-String
$planName = if ($powerPlan -match "\(([^)]+)\)\s*$") { $matches[1] } else { "Unknown" }

$usbSuspendQuery = powercfg /query SCHEME_CURRENT 2a737441-1930-4402-8d77-b2bebba308a3 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 2>&1 | Out-String
$usbSuspend = if ($usbSuspendQuery -match "Current AC Power Setting Index:\s*0x0*1") { $true } else { $false }

$fastBootReg = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power" -Name HiberbootEnabled -ErrorAction SilentlyContinue
$fastBoot = $fastBootReg.HiberbootEnabled -eq 1

# OS channel
$buildNum = [int]$os.BuildNumber
$channel = if ($buildNum -ge 26200) { "Windows_Insider_Canary" }
    elseif ($buildNum -ge 26100) { "Windows_11_24H2_Stable" }
    elseif ($buildNum -ge 22631) { "Windows_11_23H2_Stable" }
    elseif ($buildNum -ge 22621) { "Windows_11_22H2_Stable" }
    elseif ($buildNum -ge 22000) { "Windows_11_21H2_Stable" }
    elseif ($buildNum -ge 19045) { "Windows_10_22H2_Stable" }
    else { "Unknown_Build_$buildNum" }

# Device health string
$healthStr = if ($problemDevices.Count -eq 0) {
    "All OK ($($allDevices.Count) devices)"
} else {
    "$($problemDevices.Count) problems / $($allDevices.Count) devices"
}

# ==========================================
# HID DEVICES + CPU DIRECT DETECTION
# ==========================================

function Get-UsbDeviceInfo {
    param([string]$InstanceId)

    $result = @{
        ControllerType = ""  # "CPU" or "Chipset"
        ControllerInfo = $null
        HasHub = $false
        HubName = ""
        HubCount = 0
        ChipCount = 0
    }

    $currentId = $InstanceId
    $count = 0

    while ($currentId -and $count -lt 15) {
        $count++
        $dev = Get-PnpDevice -InstanceId $currentId -ErrorAction SilentlyContinue

        # Found Root Hub - next parent is the controller
        if ($currentId -match "ROOT_HUB") {
            try {
                $parentProp = Get-PnpDeviceProperty -InstanceId $currentId -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop
                $ctrlId = $parentProp.Data

                # Extract VID/DID and use database lookup
                $vid = ""; $did = ""
                if ($ctrlId -match 'VEN_([0-9A-Fa-f]{4})') { $vid = $Matches[1].ToLower() }
                if ($ctrlId -match 'DEV_([0-9A-Fa-f]{4})') { $did = $Matches[1].ToLower() }

                $ctrlInfo = Get-ControllerInfo -Vid $vid -Did $did
                $result.ControllerInfo = $ctrlInfo
                $result.ControllerType = $ctrlInfo.Type
                $result.ChipCount = $ctrlInfo.ChipCount + $result.HubCount
            } catch { }
            break
        }

        # Check for USB hub (not Root Hub)
        if ($dev -and $dev.FriendlyName -match "Hub" -and $dev.FriendlyName -notmatch "Root") {
            $result.HasHub = $true
            $result.HubName = $dev.FriendlyName
            $result.HubCount++
        }

        # Get parent
        try {
            $parentProp = Get-PnpDeviceProperty -InstanceId $currentId -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop
            $currentId = $parentProp.Data
        } catch {
            break
        }
    }

    return $result
}

# Get HID devices (mice, keyboards, gamepads)
$hidDevices = Get-PnpDevice -Class "Mouse", "Keyboard", "HIDClass" -Status OK -ErrorAction SilentlyContinue |
    Where-Object { $_.InstanceId -match "^USB\\" -or $_.InstanceId -match "^HID\\" }

$hidDevicesArray = @($hidDevices | ForEach-Object {
    # Skip generic HID sub-interfaces (but keep game controllers)
    if ($_.FriendlyName -match "^HID-compliant (mouse|keyboard|device|vendor|consumer|system)") { return }
    if ($_.FriendlyName -match "^USB Input Device$|^HID Keyboard Device$") { return }

    $instanceId = $_.InstanceId

    # For HID devices, trace up to USB parent
    $usbParent = $instanceId
    if ($instanceId -match "^HID\\") {
        try {
            $parent = Get-PnpDeviceProperty -InstanceId $instanceId -KeyName "DEVPKEY_Device_Parent" -ErrorAction Stop
            $usbParent = $parent.Data
        } catch { }
    }

    $usbInfo = Get-UsbDeviceInfo -InstanceId $usbParent

    # Extract VID/PID
    $devVid = "Unknown"
    $devPid = "Unknown"
    if ($usbParent -match "VID_([0-9A-F]{4})") { $devVid = $matches[1] }
    if ($usbParent -match "PID_([0-9A-F]{4})") { $devPid = $matches[1] }

    # Determine status based on chip count (0=CPU direct, 1=Chipset, 2+=Hub)
    $chipCount = $usbInfo.ChipCount
    $status = switch ($chipCount) {
        0 { "BEST" }           # Direct to CPU
        1 { "CHIPSET" }        # Through chipset
        default { "HUB" }      # Through hub(s)
    }
    # Add hub indicator if applicable
    if ($usbInfo.HasHub -and $chipCount -le 1) {
        $status = if ($chipCount -eq 0) { "CPU+HUB" } else { "CHIPSET+HUB" }
    }

    [ordered]@{
        name = $_.FriendlyName
        class = $_.Class
        vid = $devVid
        pid = $devPid
        controller = $usbInfo.ControllerType
        controllerInfo = $usbInfo.ControllerInfo
        chipCount = $chipCount
        hasHub = $usbInfo.HasHub
        hubName = $usbInfo.HubName
        status = $status
        instanceId = $instanceId
    }
})

# ==========================================
# DEVICE DETECTION
# ==========================================

# Known MH4 VID:PID
$mh4Vid = "054C"
$mh4Pid = "05C4"

# Try to auto-detect MH4 board
$selectedDevice = $null
foreach ($hid in $hidDevicesArray) {
    if ($hid.vid -eq $mh4Vid -and $hid.pid -eq $mh4Pid) {
        $selectedDevice = $hid
        break
    }
}

# ==========================================
# BUILD OUTPUT STRUCTURE
# ==========================================

$reportId = New-ShortId
$timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ssZ"

# Build memory DIMMs array
$slotIndex = 0
$dimmsArray = @($mem | ForEach-Object {
    [ordered]@{
        slot = $slotIndex++
        size = "$([math]::Round($_.Capacity / 1GB, 0))GB"
        brand = $_.Manufacturer
        speed = $_.Speed
    }
})

# Build USB controllers array using database lookup
$controllersArray = @($usbCtrl | ForEach-Object {
    $vid = ""; $did = ""
    if ($_.PNPDeviceID -match 'VEN_([0-9A-Fa-f]{4})') { $vid = $Matches[1].ToLower() }
    if ($_.PNPDeviceID -match 'DEV_([0-9A-Fa-f]{4})') { $did = $Matches[1].ToLower() }

    $ctrlInfo = Get-ControllerInfo -Vid $vid -Did $did

    $driverSource = if ($_.Name -match "Microsoft") { "Microsoft_Inbox" }
                    elseif ($_.Name -match "AMD|Intel|ASMedia|Renesas|VIA") { "Vendor" }
                    else { "Unknown" }
    [ordered]@{
        type = $ctrlInfo.Type
        name = $ctrlInfo.Name
        platform = $ctrlInfo.Platform
        usb = $ctrlInfo.USB
        chipCount = $ctrlInfo.ChipCount
        driver = $driverSource
    }
})

# Build storage array
$storageArray = @($disks | ForEach-Object {
    [ordered]@{
        name = $_.Model
        size = "$([math]::Round($_.Size / 1GB, 0))GB"
        firmware = $_.FirmwareRevision
    }
})

# Build the final profile
$Profile = [ordered]@{
    id = $reportId
    submitted = $timestamp
    status = $reportType
    email = $userEmail

    tested = [ordered]@{
        board = $board
        software = $softwareVersion
        cable = $cable
        port = $port
        poll = $pollRates
    }

    notes = $problemDescription

    system = [ordered]@{
        name = "$($cs.Manufacturer) $($cs.Model)"
        product = $cs.SystemFamily
        serial = $se.SerialNumber
    }

    motherboard = [ordered]@{
        board = "$($bb.Manufacturer) $($bb.Product)"
        chipset = $detectedChipset
        bios = $bios.SMBIOSBIOSVersion
    }

    cpu = [ordered]@{
        model = $cpu.Name.Trim()
        cores = "$($cpu.NumberOfCores)C / $($cpu.NumberOfLogicalProcessors)T"
        socket = $cpu.SocketDesignation
    }

    memory = [ordered]@{
        total = "$totalGB GB"
        slots = "$($mem.Count)/$($memArray.MemoryDevices)"
        dimms = $dimmsArray
    }

    usb = [ordered]@{
        ports = $usbPorts
        controllers = $controllersArray
    }

    device = if ($selectedDevice) {
        [ordered]@{
            name = $selectedDevice.name
            vid = $selectedDevice.vid
            pid = $selectedDevice.pid
            controller = $selectedDevice.controller
            hasHub = $selectedDevice.hasHub
            status = $selectedDevice.status
        }
    } else { $null }

    storage = $storageArray

    gpu = [ordered]@{
        model = $gpu.Name
        driver = $gpu.DriverVersion
    }

    deviceHealth = $healthStr

    power = [ordered]@{
        plan = $planName
        fastBoot = $fastBoot
        usbSuspend = $usbSuspend
    }

    os = [ordered]@{
        version = $os.Caption
        build = $os.BuildNumber
        channel = $channel
    }
}

# ==========================================
# HUMAN-READABLE OUTPUT (for raw field)
# ==========================================
$rawOutput = [System.Text.StringBuilder]::new()

[void]$rawOutput.AppendLine("RIG PROFILE: $($cs.Name)")
[void]$rawOutput.AppendLine("=" * 50)
[void]$rawOutput.AppendLine("ID: $reportId")
[void]$rawOutput.AppendLine("Submitted: $timestamp")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("TESTED")
[void]$rawOutput.AppendLine("  Board: $board")
[void]$rawOutput.AppendLine("  Software: $softwareVersion")
[void]$rawOutput.AppendLine("  Cable: $cable")
[void]$rawOutput.AppendLine("  Port: $port")
[void]$rawOutput.AppendLine("  Poll: 8k=$($pollRates['8000']) 4k=$($pollRates['4000']) 2k=$($pollRates['2000']) 1k=$($pollRates['1000'])")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("STATUS: $reportType")
if ($problemDescription) {
    [void]$rawOutput.AppendLine("NOTES: $problemDescription")
}
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("SYSTEM")
[void]$rawOutput.AppendLine("  $($Profile.system.name)")
[void]$rawOutput.AppendLine("  $($Profile.system.product)")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("MOTHERBOARD")
[void]$rawOutput.AppendLine("  $($Profile.motherboard.board)")
[void]$rawOutput.AppendLine("  Chipset: $($Profile.motherboard.chipset)")
[void]$rawOutput.AppendLine("  BIOS: $($Profile.motherboard.bios)")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("CPU: $($Profile.cpu.model)")
[void]$rawOutput.AppendLine("     $($Profile.cpu.cores) | $($Profile.cpu.socket)")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("MEMORY: $($Profile.memory.total) ($($Profile.memory.slots) slots)")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("USB CONTROLLERS")
foreach ($ctrl in $controllersArray) {
    [void]$rawOutput.AppendLine("  [$($ctrl.type)] $($ctrl.name) ($($ctrl.driver))")
}
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("MH4 DEVICE")
if ($selectedDevice) {
    [void]$rawOutput.AppendLine("  $($selectedDevice.name) [$($selectedDevice.status)]")
    [void]$rawOutput.AppendLine("  VID:PID $($selectedDevice.vid):$($selectedDevice.pid)")
    [void]$rawOutput.AppendLine("  Controller: $($selectedDevice.controller)")
    if ($selectedDevice.hasHub) {
        [void]$rawOutput.AppendLine("  Hub: $($selectedDevice.hubName)")
    }
} else {
    [void]$rawOutput.AppendLine("  (not detected)")
}
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("STORAGE")
foreach ($disk in $storageArray) {
    [void]$rawOutput.AppendLine("  $($disk.name) $($disk.size)")
}
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("GPU: $($Profile.gpu.model)")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("DEVICE HEALTH: $healthStr")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("POWER")
[void]$rawOutput.AppendLine("  Plan: $planName")
[void]$rawOutput.AppendLine("  Fast Boot: $fastBoot")
[void]$rawOutput.AppendLine("  USB Suspend: $usbSuspend")
[void]$rawOutput.AppendLine("")
[void]$rawOutput.AppendLine("OS: $($Profile.os.version)")
[void]$rawOutput.AppendLine("    Build $($Profile.os.build) ($channel)")

# Add raw to profile
$Profile.raw = $rawOutput.ToString()

# Prepare output path (will save after submission flow)
if ([string]::IsNullOrEmpty($OutputFileName)) {
    $OutputFileName = "$reportId.json"
}
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$OutputPath = Join-Path $ScriptDir $OutputFileName

# ==========================================
# DISPLAY TO CONSOLE
# ==========================================
Write-Host ""
Write-Host "RIG PROFILE: $reportId" -ForegroundColor $c.Header
Write-Host ("=" * 60) -ForegroundColor $c.Header
Write-Host ""

Write-Host "TESTED" -ForegroundColor $c.Label
Write-Host "  Board:    $board" -ForegroundColor $c.Value
Write-Host "  Software: $softwareVersion" -ForegroundColor $c.Value
Write-Host "  Cable:    $cable" -ForegroundColor $c.Text
Write-Host "  Port:     $port" -ForegroundColor $c.Text
$pollStr = "8k=$($pollRates['8000']) 4k=$($pollRates['4000']) 2k=$($pollRates['2000']) 1k=$($pollRates['1000'])"
Write-Host "  Poll:     $pollStr" -ForegroundColor $c.Text
Write-Host ""

# Status
$statusColor = if ($reportType -eq "working") { $c.OK } else { $c.Error }
Write-Host "STATUS: $($reportType.ToUpper())" -ForegroundColor $statusColor
if ($problemDescription) {
    Write-Host "  $problemDescription" -ForegroundColor $c.Text
}
Write-Host ""

Write-Host "SYSTEM" -ForegroundColor $c.Label
Write-Host "  $($Profile.system.name)" -ForegroundColor $c.Value
Write-Host "  $($Profile.system.product) | S/N: $($Profile.system.serial)" -ForegroundColor $c.Text
Write-Host ""

Write-Host "MOTHERBOARD" -ForegroundColor $c.Label
Write-Host "  $($Profile.motherboard.board)" -ForegroundColor $c.Value
Write-Host "  Chipset: $($Profile.motherboard.chipset) | BIOS: $($Profile.motherboard.bios)" -ForegroundColor $c.Text
Write-Host ""

Write-Host "CPU" -ForegroundColor $c.Label
Write-Host "  $($Profile.cpu.model)" -ForegroundColor $c.Value
Write-Host "  $($Profile.cpu.cores) | $($Profile.cpu.socket)" -ForegroundColor $c.Text
Write-Host ""

Write-Host "MEMORY" -ForegroundColor $c.Label
Write-Host "  $($Profile.memory.total) ($($Profile.memory.slots) slots)" -ForegroundColor $c.Value
foreach ($dimm in $dimmsArray) {
    Write-Host "  [$($dimm.slot)] $($dimm.size) $($dimm.brand) @ $($dimm.speed)MHz" -ForegroundColor $c.Text
}
Write-Host ""

Write-Host "USB CONTROLLERS" -ForegroundColor $c.Label
foreach ($ctrl in $controllersArray) {
    Write-Host "  [$($ctrl.type)] $($ctrl.name) ($($ctrl.driver))" -ForegroundColor $c.Value
}
Write-Host ""

if ($selectedDevice) {
    Write-Host "MH4 BOARD DETECTED:" -ForegroundColor $c.OK
    $statusColor = switch ($selectedDevice.status) {
        "BEST" { $c.OK }
        "CPU+HUB" { $c.Tip }
        "CHIPSET" { $c.Tip }
        "CHIPSET+HUB" { $c.Error }
        "HUB" { $c.Error }
        default { $c.Text }
    }

    Write-Host "  $($selectedDevice.name) " -ForegroundColor $c.Value -NoNewline
    Write-Host "[$($selectedDevice.status)]" -ForegroundColor $statusColor
    Write-Host "  VID:PID $($selectedDevice.vid):$($selectedDevice.pid)" -ForegroundColor $c.Text
    Write-Host "  Controller: $($selectedDevice.controller)" -ForegroundColor $c.Text
    if ($selectedDevice.controllerInfo) {
        Write-Host "  Platform: $($selectedDevice.controllerInfo.Platform)" -ForegroundColor $c.Text
    }
    if ($selectedDevice.hasHub) {
        Write-Host "  Hub: $($selectedDevice.hubName)" -ForegroundColor $c.Tip
    }
    Write-Host ""
    Write-Host "  BEST=CPU direct (0 chips) | CHIPSET=1 chip | HUB=2+ chips" -ForegroundColor $c.Text
} else {
    Write-Host "MH4 BOARD NOT DETECTED" -ForegroundColor $c.Error
}
Write-Host ""

Write-Host "STORAGE" -ForegroundColor $c.Label
foreach ($disk in $storageArray) {
    Write-Host "  $($disk.name) $($disk.size)" -ForegroundColor $c.Value
}
Write-Host ""

Write-Host "GPU" -ForegroundColor $c.Label
Write-Host "  $($Profile.gpu.model)" -ForegroundColor $c.Value
Write-Host "  Driver: $($Profile.gpu.driver)" -ForegroundColor $c.Text
Write-Host ""

Write-Host "DEVICE HEALTH" -ForegroundColor $c.Label
if ($problemDevices.Count -eq 0) {
    Write-Host "  $healthStr" -ForegroundColor $c.OK
} else {
    Write-Host "  $healthStr" -ForegroundColor $c.Error
}
Write-Host ""

Write-Host "POWER" -ForegroundColor $c.Label
Write-Host "  Plan: $planName" -ForegroundColor $c.Value
Write-Host "  Fast Boot: $fastBoot | USB Suspend: $usbSuspend" -ForegroundColor $c.Text
Write-Host ""

Write-Host "OS" -ForegroundColor $c.Label
Write-Host "  $($Profile.os.version)" -ForegroundColor $c.Value
Write-Host "  Build $($Profile.os.build) ($channel)" -ForegroundColor $c.Text
Write-Host ""

# ==========================================
# CHECK FOR KNOWN ISSUES
# ==========================================
Write-Host ("=" * 60) -ForegroundColor $c.Header
Write-Host "KNOWN ISSUES" -ForegroundColor $c.Label

# Get the first failing poll rate for the check
$failingPoll = ($pollRates.GetEnumerator() | Where-Object { $_.Value -eq "fail" } | Select-Object -First 1).Key
if (-not $failingPoll) { $failingPoll = "0" }

# Get first controller name for check
$firstController = if ($controllersArray.Count -gt 0) { $controllersArray[0].name } else { "UNKNOWN" }

$checkUrl = "https://tools.mariusheier.com/rig/check?chipset=$detectedChipset&controller=$firstController&poll=$failingPoll"
$knownIssues = $null

try {
    $checkResponse = Invoke-RestMethod -Uri $checkUrl -Method Get -TimeoutSec 10 -ErrorAction Stop
    $knownIssues = $checkResponse
} catch {
    # Silently fail - endpoint might not exist yet
}

if ($knownIssues) {
    $hasMatches = $false
    if ($knownIssues.chipset -and $knownIssues.chipset.reports -gt 0) {
        Write-Host "  * $($knownIssues.chipset.reports) reports with $($knownIssues.chipset.name) chipset" -ForegroundColor $c.Tip
        $hasMatches = $true
    }
    if ($knownIssues.controller -and $knownIssues.controller.reports -gt 0) {
        Write-Host "  * $($knownIssues.controller.reports) reports with $($knownIssues.controller.name) controller" -ForegroundColor $c.Tip
        $hasMatches = $true
    }
    if ($knownIssues.poll -and $knownIssues.poll.reports -gt 0) {
        Write-Host "  * $($knownIssues.poll.reports) reports at $($knownIssues.poll.rate) Hz" -ForegroundColor $c.Tip
        $hasMatches = $true
    }
    if (-not $hasMatches) {
        Write-Host "  No similar reports found." -ForegroundColor $c.Text
    }
    Write-Host ""
    Write-Host "Submit your profile to help improve compatibility data." -ForegroundColor $c.Text
} else {
    Write-Host "  Could not check for known issues (offline or endpoint unavailable)." -ForegroundColor $c.Text
    Write-Host ""
    Write-Host "Submit your profile to help identify new issues." -ForegroundColor $c.Text
}

Write-Host ("=" * 60) -ForegroundColor $c.Header
Write-Host ""

# ==========================================
# SUBMISSION FLOW
# ==========================================

# Optional email
Write-Host "Email (to notify if a fix is found):" -ForegroundColor $c.Label
Write-Host "Optional, press Enter to skip" -ForegroundColor $c.Text
$emailInput = Read-Host
if (-not [string]::IsNullOrWhiteSpace($emailInput)) {
    $userEmail = $emailInput
    $Profile.email = $userEmail
}

# Submit prompt
Write-Host ""
$submitChoice = Read-Host "Publish to help others? (Y/N)"

if ($submitChoice -match "^[Yy]") {
    # Update the profile with email before sending
    $Profile.email = $userEmail
    $jsonPayload = $Profile | ConvertTo-Json -Depth 6 -Compress

    $submitUrl = "https://tools.mariusheier.com/rig"
    $submitted = $false

    try {
        $response = Invoke-RestMethod -Uri $submitUrl -Method Post -Body $jsonPayload -ContentType "application/json" -TimeoutSec 15 -ErrorAction Stop
        $submitted = $true
        Write-Host ""
        Write-Host "Submitted!" -ForegroundColor $c.OK
        Write-Host "View your report: https://tools.mariusheier.com/rig/$reportId" -ForegroundColor $c.Value
    } catch {
        Write-Host ""
        Write-Host "Could not submit (network error)." -ForegroundColor $c.Error
    }
} else {
    Write-Host ""
    Write-Host "Not submitted." -ForegroundColor $c.Text
}

# Save JSON locally (update with email if provided)
$Profile.email = $userEmail
$jsonOutput = $Profile | ConvertTo-Json -Depth 6 -Compress:$false
$jsonOutput | Out-File -FilePath $OutputPath -Encoding UTF8 -Force

Write-Host ""
$fileSize = (Get-Item $OutputPath).Length
Write-Host "Saved: $OutputPath ($([math]::Round($fileSize/1024, 1)) KB)" -ForegroundColor $c.Text
Write-Host ""
