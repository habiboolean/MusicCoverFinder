import datetime
import json
import os
import re
import time
import urllib.parse
from http.cookies import SimpleCookie
import http.client as httplib

import openpyxl
import requests
from PIL import Image, UnidentifiedImageError
from bs4 import BeautifulSoup
from openpyxl.worksheet.worksheet import Worksheet
from urllib3.exceptions import ReadTimeoutError
from rich import print

raw_cookie = os.getenv('COOKIES')
header = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36"}
request_delay = 2   # delay to not get banned by yandex (less than 2 sec not recommended)
column_index = 0    # column in excel file to read names of songs from


def have_internet() -> bool:
    """
    Checks if internet connection is ON

    :return: True if it can read headers of google DNS server (8.8.8.8)
    """
    conn = httplib.HTTPSConnection("8.8.8.8", timeout=5)
    try:
        conn.request("HEAD", "/")
        return True
    except Exception:
        return False
    finally:
        conn.close()


def get_excel_sheet(filename: str) -> Worksheet:
    """
    Get active worksheet

    :param filename: path to the file
    :return: worksheet
    """
    wb = openpyxl.load_workbook(filename)
    return wb.active


def strfdelta(tdelta, fmt):
    """
    Format timedelta

    :param tdelta: timedelta
    :param fmt: format i.e. {hours}:{minutes}:{seconds}
    """
    d = {"days": tdelta.days}
    d["hours"], rem = divmod(tdelta.seconds, 3600)
    d["minutes"], d["seconds"] = divmod(rem, 60)
    return fmt.format(**d)


def print_status(start_time: datetime, current_position: int, total_positions: int, file_name: str, status: str):
    """
    Prints status line, i.e.:

    SUCCESS:  Alexander et Son Orch. - Pour Toi, Rio-Rita

    288 out of 14372. estimated time left 1:5:11. Time per single file in avg 3 sec

    :param start_time: datetime when operation started
    :param current_position: current line in excel book
    :param total_positions: total lines in excel book
    :param file_name: name of current song we searching for
    :param status: string to print in a first place. i.e. (SUCCESS: ERROR:)
    """
    current_position += 1

    time_diff = datetime.datetime.now() - start_time
    speed = time_diff.seconds / current_position
    left = (total_positions - current_position) * speed
    formatted_time_left = strfdelta(datetime.timedelta(seconds=left), "{hours}:{minutes}:{seconds}")

    print(f'\r{status} {file_name}')
    print("\r" + f"{current_position} out of {total_positions}. estimated time left {formatted_time_left}. Time per single file in avg {speed} sec",
          end="")


def save_image(url_hires: str, url_lowres: str, file_name: str, file_format: str = 'PNG') -> bool:
    """
    Trying to save image from given URL's

    :param url_hires: link to HIGH resolution image
    :param url_lowres: link to LOW resolution image
    :param file_name: name file for saving
    :param file_format: i.e.: PNG, JPG, WEBM...
    :raises UnidentifiedImageError: if PIL cant read image
    :return: True if saved successfully, else - False
    """
    if url_hires.startswith('//'):
        url_hires = 'https:' + url_hires

    if url_lowres.startswith('//'):
        url_lowres = 'https:' + url_lowres

    try:
        if url_lowres == url_hires:
            img = Image.open(requests.get(url_lowres, stream=True, timeout=15).raw)
            img.save(f'img/lo-res/{file_name}.{file_format}', format=file_format)
            img.save(f'img/hi-res/{file_name}.{file_format}', format=file_format)
        else:
            img = Image.open(requests.get(url_lowres, stream=True, timeout=15).raw)
            img.save(f'img/lo-res/{file_name}.{file_format}', format=file_format)
            img = Image.open(requests.get(url_hires, stream=True, timeout=15).raw)
            img.save(f'img/hi-res/{file_name}.{file_format}', format=file_format)
    except UnidentifiedImageError as e:
        print(f'\nERROR: Cannot save image for "{file_name}.{file_format}"')
        return False
    except Exception as e:
        print(f'\nERROR: {e} for "{file_name}.{file_format}"')
        return False

    return True


def create_img_folders():
    """
    Creates folders for images to save if they not exist
    """
    path = 'img/hi-res'
    os.makedirs(path, exist_ok=True)
    path = 'img/lo-res'
    os.makedirs(path, exist_ok=True)


def main():
    create_img_folders()

    sheet = get_excel_sheet('For Tigrik2.xlsx')
    total = sheet.max_row

    start_time = datetime.datetime.now()  # takes current time on start to measure average time to complete operations

    for index, row in enumerate(sheet.iter_rows()):
        x = row[column_index].value

        # check if file already exist to not redownload it
        if os.path.exists(f'img/hi-res/{x}.png'):
            continue

        x = re.sub(r"[\(\[].*?[\)\]]", "", x)
        x = x.strip()
        search_query = x + ' обложка'

        no_conn_counter = 0
        while not have_internet():
            no_conn_counter += 1
            print(f'\rNO INTERNET: {no_conn_counter}', end="")
            time.sleep(10)

        try:
            cookie = SimpleCookie()
            cookie.load(raw_cookie)
            cookies = {k: v.value for k, v in cookie.items()}
            request = requests.get(f'https://yandex.ru/images/search?iorient=square&text='
                                   f'{urllib.parse.quote(search_query)}', headers=header, cookies=cookies, timeout=10)
        except ReadTimeoutError as e:
            print_status(start_time, index, total, row[column_index].value, 'ERROR: URL TIMEOUT for ')
            continue
        except Exception as e:
            print_status(start_time, index, total, row[column_index].value, f'ERROR: request {e} ')
            continue

        if not request:
            print_status(start_time, index, total, row[column_index].value, 'ERROR: request is NONE for ')
            continue

        time.sleep(request_delay)

        soup = BeautifulSoup(request.content, "html.parser")

        item = soup.find(class_='serp-item_type_search')
        try:
            image_json = json.loads(item['data-bem'])
        except:
            print_status(start_time, index, total, row[column_index].value, 'ERROR: cant resolve data-bem for ')
            continue

        hi_res_url = image_json['serp-item']['preview'][0]['url']

        item = soup.find(class_='serp-item__link')
        low_res_url = item.find('img')['src']

        if low_res_url == '':
            print_status(start_time, index, total, row[column_index].value, 'ERROR: URL was empty for ')
            continue

        if save_image(hi_res_url, low_res_url, row[column_index].value):
            print_status(start_time, index, total, row[column_index].value, 'SUCCESS: ')
        else:
            print_status(start_time, index, total, row[column_index].value, 'FAILED: ')


if __name__ == '__main__':
    main()
