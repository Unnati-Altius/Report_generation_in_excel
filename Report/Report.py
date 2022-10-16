class RfpReport:
    def __init__(self, worksheet, company, services, mapping):
        """
            input: 
                worksheet: This is the instance of excel file
                company: List of all the companies
                services: List of all the services
                mapping: {company: [[service1, price], [service2: price]...]}
                         mapping of companies with its services and price
            output:
                None
        """

        self.worksheet = worksheet
        self.company = company
        self.services = services
        self.mapping = mapping

    def add_companies(self):
        """
            Desc: This function adds companies name 
            as a header in the excel file
        """

        for i in range(len(self.company)):
            self.worksheet.write(0, i+1, self.company[i])


    def add_services(self):
        """
            Desc: This function adds services name 
            as a header in the excel file
        """

        for i in range(len(self.services)):
            self.worksheet.write(i+2, 0, self.services[i])
            self.mark_services(i)

    def mark_services(self, service_idx):
        """
            Desc: This function maps services name with
            companies in the excel file
        """

        for j, val in enumerate(self.company):
            s = self.mapping[val]
            se = []
            for v in s:
                se.append(v[0])
            if self.services[service_idx] in se:
                self.worksheet.write(service_idx+2, j+1, "yes")
            else:
                self.worksheet.write(service_idx+2, j+1, "no")
