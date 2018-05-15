class Cylinder:
    'Defines a gas cylinder (or dewar).'

    def __init__(self, size, gas="", hydrotest="Unknown"):
        self.size = size
        self.gas = gas
        self.hydrotest = hydrotest


class CO2Air(Cylinder):
    'Defines a CO2-Air cylinder.'

    def __init__(self, SN, Lot, CO2, O2, N2):
        Cylinder.__init__(self, 'H', "CO2-Air")
        self.SN = SN
        self.Lot = Lot
        self.CO2 = CO2
        self.O2 = O2
        self.N2 = N2

    def __str__(self):
        return f"SN: {self.SN}, Lot: {self.Lot}, CO2: {self.CO2}, \
O2: {self.O2}, N2: {self.N2}"


class Nitrogen(Cylinder):
    'Defines a Liquid nitrogen dewar, or cylinder.'

    def __init__(self, SN, Lot, N2, O2="ND", CO="ND"):
        Cylinder.__init__(self, 'LS240', "N2")
        self.SN = SN
        self.Lot = Lot
        self.N2 = N2
        self.O2 = O2
        self.CO = CO

    def __str__(self):
        return f"SN: {self.SN}, Lot: {self.Lot}, N2: {self.N2}, O2: {self.O2}, CO: {self.CO}"