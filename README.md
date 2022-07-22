
# Music-Cover-Finder :musical_note:

- Task from customer was: Have Excel file with thousands (around 15k) album/songs names. We need a cover picture for each of them, preferably in low-res (less than 500x500px) and hi-res (1000x1000px or more).
- After research, found that names was in russian and english. Some of albums was super rare, and couldnt be found even at spotify like resources. That is reason why only yandex was used, since it gave best results and was able to find nearly perfect matches.

# DEMO

- Sample data:

<img src="https://user-images.githubusercontent.com/105993976/180548093-8bf97d9c-ba8f-448b-83a9-fe03d961a217.png" height="500">

- Example of work

https://user-images.githubusercontent.com/105993976/180547584-c9f6d5ee-0e85-4e3a-89b2-1838f5922c32.mp4


## Tech Stack
- **Python** 
- [**Beautiful Soup**](https://www.crummy.com/software/BeautifulSoup/bs4/doc/)
- [**openpyxl**](https://openpyxl.readthedocs.io/en/stable/)
- [**Rich**](https://github.com/Textualize/rich)
- [**Pillow**](https://github.com/python-pillow/Pillow)
- [**Requests**](https://github.com/psf/requests)

## Features:
- Uses yandex image search to get best match for excel data
- Save pics in differen formats (webp, png, jpg ... basicaly any file pillow supoort)
- Save a preview size (320x320) and highest available for same search result
- Shows user status, calculates time before finish and average time to process single request
- Show logs for each request (SUCCESS if file was saved, or ERROR if coundt process file)
- Detects if internet connection lost and wait till it back. (fast method, thanks to [lvelin and boatcoder (stackoverflow)](https://stackoverflow.com/a/29854274)
- If process was terminated (any reason) it will continue from the start, but skipping if file already saved.

## TODO:
- Cover with tests (pytest)
- Extracting album name to have option to save file with album name
- Check if same picture already been saved before (by link, or maybe openCV similarity comparsion)
- Limit hi-res file size
- Notification to Telegram when finish process
- Use selenium for automatic cookies generation (this part not sure, used Bs4 due to its more familiar and i had no time limits mostly for the task)

## Quick Start

- Fork and Clone the repository using:
```
git clone https://github.com/habiboolean/MusicCoverFinder.git
```
- Install dependencies using:
```
pip install -r requirements.txt
```
- Run main.py:
```
python main.py
```

>![forthebadge made-with-python](http://ForTheBadge.com/images/badges/made-with-python.svg)


