import os
import requests
from bs4 import BeautifulSoup
import xlsxwriter

class Movie:
	ismovie=False
	title="ABC"
	year="0"
	genre=""
	rating="0.0"
	part="0"
	error=False
	youtubelink=""

def getmovieinfo(name):
	s=requests.session()
	guessurl= "http://guessit.io/guess?filename="+name
	#print guessurl
	movieguesspage=s.get(guessurl)
	movieinfo=movieguesspage.content

	mov=Movie()

	'''if "\"type\"" in movieinfo:
		type_start=movieinfo.find("\"type\"")+9
		type_end=movieinfo[type_start:].find("\"")+type_start
		movietype=movieinfo[type_start:type_end]
		#print movietype
		if(movietype=="movie"):
			mov.ismovie=True
		else:
			return mov'''
	#print movieinfo
	if "title" in movieinfo and "500 Internal Server Error" not in movieinfo:
		title_start=movieinfo.find("title")+9
		title_end=movieinfo[title_start:].find("\"")+title_start
		mov.title=movieinfo[title_start:title_end]
		#print mov.title
	else:
		mov.error=True
		mov.title=name
		#print "1"
		return mov

	if "year" in movieinfo:
		year_start=movieinfo.find("year")+7
		year_end=movieinfo[year_start:].find(",")+year_start
		mov.year=movieinfo[year_start:year_end]
	#print mov.year

	if "part" in movieinfo:
		part_start=movieinfo.find("\"part\"")+8
		part_end=movieinfo[part_start:].find(",")+part_start
		mov.part=movieinfo[part_start:part_end]
	#print movieinfo
	
	#print mov.part
	#print "2"
	return mov

def getrating(mov):
	s=requests.session()
	if mov.error==True:
		url = "https://www.google.co.in/search?q="+mov.title+" movie imdb"
	#print mov.title
	else:
		url = "https://www.google.co.in/search?q="+mov.title
		if mov.part != "0":
			url = url + " part " + mov.part
		if mov.year != "0":
			url = url + " year "+mov.year
		url = url + " imdb"
	#print url
	page = s.get(url)
	soup = BeautifulSoup(page.content)
	a_s=soup.find_all('a')
	mov.error=True
	for a in a_s:
		link=a.get("href")
		if "http://www.imdb.com/title/tt" in link:	
			mov.error=False
			#print(link)
			start=link.find("http")
			end=link.find("title/tt")
			end=end+15;
			#end=link[end:].find("/")+end
			completelink=link[start:end]
			#print(completelink)
			newpage=s.get(completelink)
			newsoup=BeautifulSoup(newpage.content)
			heading=newsoup.find('h1')
			mov.title=heading.find('span',attrs={'itemprop':'name'}).string
			if heading.find('a')==None:
				yearstring=heading.find('span',attrs={'class':'nobr'}).string
				mov.year=yearsting[1:-1]
			else:
				yearstring=heading.find('a').get('href')
				mov.year=yearstring[6:10]
			#print mov.year
	
			rating_tags=newsoup.find('span', attrs={'itemprop':'ratingValue'})
			if(rating_tags == None):
				mov.error=True
				return mov
			mov.rating=rating_tags.string
				
			genre_tags=newsoup.find_all('span',attrs={'itemprop':'genre'})
			for genre_tag in genre_tags:
				mov.genre=mov.genre+genre_tag.string+'|'
			mov.genre=mov.genre[:-1]
			break
	#print mov.rating
	return mov

def getyoutubelink(mov):
	s=requests.session()
	searchurl="https://www.youtube.com/results?search_query="+mov.title+"official trailer"
	page = s.get(searchurl)
	soup = BeautifulSoup(page.content)
	h_s=soup.find('h3',attrs={'class':'yt-lockup-title'})
	link = h_s.find('a').get('href')
	mov.youtubelink="https://www.youtube.com"+link
	return mov

videoExtensions = [".avi",".mp4",".mkv",".mpg",".mpeg",".mov",".wmv",".flv",".3gp",".MP4",".AVI",".WMV",".MOV",".FLV",".MKV"]
errorfiles=["nil"]
directory="F:\movies"

workbook = xlsxwriter.Workbook("Movie_Rating.xlsx")
worksheet = workbook.add_worksheet('rating')

bold = workbook.add_format({'bold': True})

title_format = workbook.add_format({'underline':  1})

rating_format = workbook.add_format()
rating_format.set_num_format('0x0F')
rating_format.set_align('right')

year_format = workbook.add_format()
year_format.set_num_format('0000')
year_format.set_align('right')

url_format = workbook.add_format({'font_color': 'blue','underline':  1})

worksheet.write('A1','Movie',bold)
worksheet.write('B1','Year',bold)
worksheet.write('C1','Rating',bold)
worksheet.write('D1','Genre',bold)
worksheet.write('E1','Trailer',bold)
worksheet.set_column(0,0,40)
#worksheet.set_column(0,0,30)
#worksheet.set_column(0,0,30)
worksheet.set_column(3,4,50)
#worksheet.set_column(4,4,50)
row=1

cur_movie = Movie()
f=open('F:/movies/error.txt','w')
#f.write('Movie,IMDb Rating,Genre')
for root, dirs, files in os.walk(directory, topdown=False):
	for name in files:
		for extension in videoExtensions:
			if name.endswith(extension):
				if os.path.getsize(os.path.join(root,name))>50000000:
					cur_movie = getmovieinfo(name)
					if cur_movie.error==True:
						f.write(os.path.join(root, name))
					cur_movie = getrating(cur_movie)
					cur_movie = getyoutubelink(cur_movie)
					if cur_movie.error==True:
						f.write(os.path.join(root, name))
					#print os.path.join(root, name)
					else:
						worksheet.write_url(row,0,os.path.join(root, name),title_format,cur_movie.title)
						worksheet.write(row,1,cur_movie.year,year_format)
						if float(cur_movie.rating)+0.1 >8.0:
							rating_format.set_font_color('green')
						else:
							rating_format.set_font_color('red')
						worksheet.write(row,2,cur_movie.rating,rating_format)
						worksheet.write(row,3,cur_movie.genre)
						worksheet.write_url(row,4,cur_movie.youtubelink,url_format)
						row += 1
						print ("completed {title}".format(title=cur_movie.title))
					#f.write('\n{title},{rating},{genre}'.format(title=cur_movie.title,rating=cur_movie.rating,genre=cur_movie.genre))
					#print('\n{title},{rating},{genre}'.format(title=cur_movie.title,rating=cur_movie.rating,genre=cur_movie.genre))
f.close()
workbook.close()
print ('Ratings assigned successfully')				
				#count=count+1
				#print(os.path.join(root, name))

#print count
