import os
filepath=os.getcwd()+'/StremCrawlingResult.xlsm'      #실행 파일과 같은 폴더에 있는 StremCrawlingResult.xlsm을 불러옴

if os.path.isfile(filepath)==True:
    import RestPriceCrawler
    RestPriceCrawler

else:
    import CatalogCrawler
    CatalogCrawler
    import PriceCrawler
    PriceCrawler
