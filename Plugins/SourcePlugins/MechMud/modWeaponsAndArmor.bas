Attribute VB_Name = "modWeaponsAndArmor"
Option Explicit

Public Enum enumWeaponDmg_template
    SmallBullet
    MediumBullet
    HeavyBullet
    ParticleSmall
    ParticleMedium
    ParticleHeavy
    RocketSmall
    RocketLarge
    RocketExtreme
    SniperBullet
    SniperBulletExtreme
End Enum

Public Enum enumPersonalWeapons_template
    'Small - no speed modifier or turn modifier
    Knife
    VibroBlade
    Hands
    RiotGun
    ShotGun
    SubMachineGun
    SubMachineGunSilenced
    PlasmaRifle
    'Large - small turn modifier
    Minigun
    PortablePPC
    PortableStinger
    PortableAntiAir
    PortableAntiTank
    PortablePlasmaCannon
    'Huge - large turn modifier
    Mortar90mm
    Mortar120MM
    Mortar250MM
    'Specials
    SalvagedPPCCannon
    SalvagedLightLaser
    SalvagedMediumLaser
    SalvagedHeavyLaser
    SalvagedSpecial
    SalvagedRocketLauncherSmall
    SalvagedRocketLauncherMedium
    SalvagedRocketLauncherLarge
End Enum

Public Enum enumPersonalShield_template
    BodyArmorLight
    BodyArmorMedium
    BodyArmorHeavy
    RiotShieldSmall
    RiotShieldBig
    UnitArmor
    UnitArmorSpecial
    PlasmaShield
    PlasmaShieldExtreme
    PlasmaShieldGodLike
End Enum

Public Enum enumMechWeapons_template
    PPCCannon
    LightLaser
    MediumLaser
    HeavyLaser
    Special
    MachineGun
    RocketLauncherSmall
    RocketLauncherMedium
    RocketLauncherLarge
    Mortar120MM
    Mortar250MM
    RocketLauncherGuidingSmall
    RocketLauncherGuidingMedium
    RocketLauncherGuidingLarge
End Enum

Public Enum enumMechShield_template
    UnionPlatingFibre
    UnionPlatingSteel
    UnionPlatingFireSteel
    UnionPlatingColdSteel
    UnionPlatingSpecial
    SkullPlatingFibre
    SkullPlatingSteel
    SkullPlatingIceSteel
    InnerSphereFibre
    InnerSphereSteel
    InnerSphereFire
    InnerSphereCold
    InnerSphereIce
    InnerSphereEarth
    ProtoTypePlasma
    ProtoTypeCloak
    ProtoTypeSpecial
    SalvageArmorLight
    SalvageArmorHeavy
End Enum


