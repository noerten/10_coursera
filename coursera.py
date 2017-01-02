import json
import random
import re
import urllib.request
from xml.etree import ElementTree

from bs4 import BeautifulSoup
from openpyxl import Workbook


COURSES_XML_URL = 'https://www.coursera.org/sitemap~www~courses.xml'
DEFAULT_FILEPATH = 'courses.xlsx'


def get_url_content(url):
    response = urllib.request.Request(url)
    return urllib.request.urlopen(response).read()


def get_courses_list(xml_response, course_quantity=20):
    xml_courses = ElementTree.fromstring(xml_response)
    random_courses = [random.choice(xml_courses)[0].text for i
                      in range(course_quantity)]
    return random_courses


def get_course_start_date(soup):
    start_date_script_tag = soup.find('script', attrs={
            'type': 'application/ld+json'
            })
    if start_date_script_tag is not None:
        start_date_dict = json.loads(start_date_script_tag.string)
        try:
            start_date = start_date_dict['hasCourseInstance'][0]['startDate']
            return start_date
        except KeyError:
            return None


def get_number_of_weeks(soup):
    number_of_weeks_tag_list = soup.find_all('div', class_='week')
    if number_of_weeks_tag_list:
        number_of_weeks_str = number_of_weeks_tag_list[-1].div.get_text()
        return int(number_of_weeks_str.split()[-1])


def get_average_rating(soup):
    average_rating_tag = soup.find('div', class_='ratings-text')
    if average_rating_tag is not None:
        average_rating_str = average_rating_tag.get_text()
        return float(re.findall('\d+\.?\d*', average_rating_str)[0])


def get_course_info(html):
    soup = BeautifulSoup(html, 'html.parser')
    title = soup.find('div', class_=['title', 'display-3-text']).get_text()
    language = (soup.find('div', class_='language-info').get_text().split()[0].
                rstrip(','))
    start_date = get_course_start_date(soup)
    number_of_weeks = get_number_of_weeks(soup)
    average_rating = get_average_rating(soup)
    return [title, language, start_date, number_of_weeks, average_rating]


def output_courses_info_to_xlsx(all_courses_info, filepath):
    header = ['Title', 'Language', 'Start Date', 'Number of Weeks',
              'Average Rating', 'Link']
    wbook = Workbook()
    wsheet = wbook.active
    wsheet.append(header)
    for course in all_courses_info:
        wsheet.append(course)
    wbook.save(filepath)


def get_filepath():
    filepath = input("Enter file path/name where to save or press 'Enter' to "
                     "save in script folder to 'courses.xlsx': ")
    if not filepath:
        filepath = DEFAULT_FILEPATH
    return filepath


def main():
    filepath = get_filepath()
    xml_response = get_url_content(COURSES_XML_URL)
    random_course_urls = get_courses_list(xml_response)
    all_courses_info = []
    for course_url in random_course_urls:
        print('Scraping', course_url)
        html = get_url_content(course_url)
        course_info = get_course_info(html)
        course_info.append(course_url)
        all_courses_info.append(course_info)
    output_courses_info_to_xlsx(all_courses_info, filepath)
    print('File saved!')


if __name__ == '__main__':
    main()
