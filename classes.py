class Info:
    def __init__(
        self,
        Date,
        store,
        TotalInterceptions,
        TotalSales,
        shortage,
        BAName,
        BAInstance,
        bits,
    ):
        self.Date = Date
        self.store = store
        self.TotalInterceptions = TotalInterceptions
        self.TotalSales = TotalSales
        self.shortage = shortage
        self.BAWhatsapp = BA(BAName, BAInstance)
        self.bits = bits

    def print(self):
        print("Date: ", self.Date)
        print("store: ", self.store)
        print("Total Interceptions: ", self.TotalInterceptions)
        print("total Sales: ", self.TotalSales)
        print("shortage: ", self.shortage)
        self.BAWhatsapp.print()


class Excel:
    def __init__(self, sheetName, store, BAName, BAInstances):
        self.sheetName = sheetName
        self.store = store
        self.BADetails = BA(BAName, BAInstances)

    def print(self):
        print("sheet name: ", self.sheetName)
        print("store: ", self.store)
        self.BADetails.print()


class BA:
    def __init__(self, name, instances):
        self.name = name
        self.instances = instances

    def print(self):
        print("BA name: ", self.name)
        print("instances: ", self.instances)
        print("-----------------")


class RedSyrup:
    def __init__(self, small, medium, large, smallSF):
        self.small = small
        self.medium = medium
        self.large = large
        self.smallSF = smallSF


class GreenSyrup:
    def __init__(self, Sandaleen, Bazooreen, Ilacheen):
        self.Sandaleen = Sandaleen
        self.Bazooreen = Bazooreen
        self.Ilacheen = Ilacheen
