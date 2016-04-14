#Result scrapper

scraper for UPTU btech result .


# You can use this application on Windows as well as Ubuntu

1. After installing Ubuntu 14.04, refresh your apt package index
    
   sudo apt-get update

2. Now, install pip using Python Version 2
   
   sudo apt-get install python-pip

3. After installing pip , install the dependencies

   pip install -r requirements.txt

4. Now, run the following git clone(specify a directory)
   
  https://github.com/JINDALG/uptu-result.git uptu

5. Firefox is required for this. Download Firefox.

6. Move to the directory in which you cloned the git repository.

   cd uptu

7. You are all set to run the development server

   scrapy crawl btech
