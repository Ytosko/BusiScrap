from modules.scraper_web import scrape
from modules.cliargs import parse_cliargs

if __name__ == "__main__":
    args = parse_cliargs()
    scrape(args)


