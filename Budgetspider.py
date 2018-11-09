import pandas as pd
import scrapy


# This spider crawls the budget document library of the Ministry of Finance in Uganda. It can help keep track the
# completeness of the library

class BudgetLibrarySpider(scrapy.Spider):
    name = "Budgetspider"
    # Start from the homepage
    start_urls = ['http://budget.go.ug']

    # main function that controls requests
    def parse(self, response):
        # gather list of URLs to the different years
        urls = response.xpath('//li/a[contains(@href, "budget/library")][contains(.//text(), "FY")]//@href').extract()
        # follow the links gathered and parse the response
        for url in urls:
            yield scrapy.Request(url=response.urljoin(url), callback=self.parse_details)

    # function that controls parsing
    def parse_details(self, response):
        # initiate dict that collects dataframes to write to excel
        dataframe_dict = {}
        # start counter
        level_count = 0
        # extract year from the current page
        year = response.css('section>div>div>div>table a::text')[0].extract().replace('/', '_')

        # loop through the levels, get the level name
        for level in response.css('div.pane-content>div>div>div>div'):
            level_name = response.css('div.pane-content>div>ul>li>a::text')[level_count].extract()
            # initialise category count
            category_count = 0

            # loop through the document category, get the category name
            for document_category in level.css('div>div>div.quicktabs-tabpage>:not(section)'):
                category_name = level.css('ul>li>a::text')[category_count].extract()
                # initialise sub-category count
                sub_category_count = 0
                # check whether there is a table that is nested in this level and parse it
                if len(document_category.css('div.view-LGPAs>div>div>table').extract()) != 0:
                    try:
                        # convert the table into a pandas dataframe, and make a dictionary out of it. Do this only if
                        # the table actually exists, else pass
                        dataframe_dict[year + '_' + str(level_count + 1) + '_' + level_name + '_' + str(
                            category_count + 1) + '_' + category_name] = pd.read_html(
                            document_category.css('div.view-LGPAs>div>div>table')[0].extract())
                    except:
                        IndexError
                else:

                    # loop through the sub category, get the sub_category name and extract tables at that level
                    for document_sub_category in document_category.css('h3 + div'):
                        sub_category_name = document_category.css('h3>a::text')[sub_category_count].extract()
                        # check whether there are tables that are nested in this level and parse them
                        if len(document_sub_category.css('div>div.quicktabs-tabpage').extract()) == 0:
                            try:
                                # convert the table into a pandas dataframe, and make a dictionary out of it. Do this only if
                                # the table actually exists, else pass
                                dataframe_dict[year + '_' + str(level_count + 1) + '_' + level_name + '_' +
                                               str(category_count + 1) + '_' + category_name + '_' +
                                               str(sub_category_count + 1) + '_' + sub_category_name] = \
                                    pd.read_html(document_sub_category.css('section table')[0].extract())
                            except: IndexError

                        else:
                            # initialise sub-sub-category count
                            sub_sub_category_count = 0
                            # loop through the sub category, get the sub_category name and extract tables at that level
                            for document_sub_sub_category in document_sub_category.css('div>div.quicktabs-tabpage'):
                                sub_sub_category_name = document_sub_category.css('ul>li>a::text')[sub_sub_category_count].extract()
                                try:
                                    # convert the table into a pandas dataframe, and make a dictionary out of it. Do this only if
                                    # the table actually exists, else pass
                                    dataframe_dict[year + '_' + str(level_count + 1) + '_' + level_name + '_' + str(
                                        category_count + 1) + '_' + category_name + '_' + str(sub_category_count + 1) + '_' +
                                        sub_category_name + '_' + str(sub_sub_category_count + 1) + '_' + sub_sub_category_name] = pd.read_html(
                                        document_sub_sub_category.css('section table')[0].extract())
                                except:
                                    IndexError

                                sub_sub_category_count += 1
                        sub_category_count += 1
                category_count += 1
            level_count += 1


        # export the dataframes to excel
        for df in dataframe_dict:
            # start excel writer
            writer = pd.ExcelWriter('./'+ df + '.xlsx', engine='xlsxwriter')
            dataframe_dict[df][0].to_excel(writer, 'Sheet1')
            writer.save()

# To do: (1) handle case where one cell holds more than one doc, (2) merge excels, (3) let scrapy check whether links to docs actually work