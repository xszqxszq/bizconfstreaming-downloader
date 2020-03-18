import os
import sys
import docx
import time
import requests
import threading

def getColor(name):
	if name.lower() == 'info':
		return '\033[1;33;320m'
	elif name.lower() == 'error':
		return '\033[1;31;320m'

def printLog(text='', resetColor='\033[0m\n', type='INFO', back=False):
	if back:
		resetColor = '\033[0m\r'
	print(getColor(type), '[{0}] '.format(type.upper()), time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()), ' ', text, sep='', end=resetColor)

def getFileExt(url):
	if '.' in url:
		return url.split('.')[-1]
	else:
		return ''

def getTsList(m3u8='', prefix=False):
	tsList, tsUrlPrefix = [], os.path.split(m3u8)[0] + '/'
	m3u8content = requests.get(m3u8).text.strip()
	lines = m3u8content.split('\r\n')
	if len(lines) < 3:
		lines = m3u8content.split('\n')
	for index, line in enumerate(lines):
		if '#EXTINF' in line:
			tsList.append(lines[index + 1])
	if prefix:
		tsList = [tsUrlPrefix + i for i in tsList]
		return tsList
	return tsList, tsUrlPrefix


def getEstimatedTsSize(m3u8):
	size, tsList = 0, getTsList(m3u8, True)
	if len(tsList) < 4:
		for i in tsList:
			size += getFileSize(i)
	else:
		size = (getFileSize(tsList[0]) + getFileSize(tsList[1]
													 ) + getFileSize(tsList[2])) / 3 * len(tsList)
	return size


def getFileSize(url):
	size = 0
	if getFileExt(url) == 'm3u8':
		size = getEstimatedTsSize(url)
	else:
		h = requests.head(url).headers
		if 'content-length' in h:
			size = int(h['content-length'])
	return size


def getVideoInfo(pageurl):
	infoapi = 'http://webcasting.bizconfstreaming.com/videos/api/get_video/{0}'
	videopage = requests.get(pageurl).text
	m_id = videopage.split('trailer: {"m_id":"')[1].split('","')[0]
	videoinfo = requests.get(infoapi.format(m_id)).json()
	if videoinfo['errcode'] != 1000:
		print('\033[1;31;320m获取失败\033[0m')
		exit()
	return videoinfo


def getBestQualityLink(playset):
	return max(playset, key=lambda i: i['resolution'])['url']


def doDownloadMp4(filename, fileurl):
	os.system(
		'aria2c -o "{0}" -j 32 "{1}" > /dev/null'.format(filename, fileurl))


def doWatchAndDownload(url):
	r = requests.get(url)
	if '活动已关闭' in r.text:
		return False
	else:
		try:
			videoInfo = getVideoInfo(url)
		except Exception:
			return False
		print()
		filename, fileurl = videoInfo['video']['name'], min([videoInfo['video']['url'],
															 getBestQualityLink(videoInfo['video']['play_set'])],
															key=lambda i: getFileSize(i))
		printLog('正在下载中…… 预计文件大小为 {0}MB'.format(
			getFileSize(fileurl) // 1048576))
		if getFileExt(fileurl) == 'm3u8':
			if os.path.exists(filename+'.mp4'):
				printLog('文件存在，已跳过')
				return True
			tsList, tsUrlPrefix = getTsList(fileurl)

			if not os.path.exists('temp'):
				os.mkdir('temp')
			with open('temp/filelist.txt', 'w') as f:
				for ts in tsList:
					f.write(tsUrlPrefix + ts + '\n')
			os.system(
				'aria2c --max-concurrent-downloads=32 --input-file=temp/filelist.txt -d temp > /dev/null')

			printLog('合并文件中……')
			final = open(filename + '.ts', 'wb')
			for ts in tsList:
				with open('temp/' + ts, 'rb') as f:
					while True:
						buf = f.read(4096)
						if buf:
							final.write(buf)
						else:
							break
				os.remove('temp/' + ts)

			printLog('合并完毕。正在转为mp4……')
			os.system(
				'ffmpeg -i "{0}.ts" -map 0 -c copy "{0}.mp4"'.format(filename))

			os.remove(filename + '.ts')
			os.remove('temp/filelist.txt')
			os.rmdir('temp')
		else:
			if '.mp4' not in filename:
				filename += '.mp4'
			if os.path.exists(filename):
				printLog('文件存在，已跳过')
				return True
			doDownloadMp4(filename, fileurl)
		printLog('下载完成。')
		return True

if __name__ == '__main__':
	if len(sys.argv) < 2:
		printLog('使用方法1：{0} http://webcasting.bizconfstreaming.com/watch/课程ID'.format(sys.argv[0]), type='error')
		printLog('使用方法2：{0} 课程链接合集.docx'.format(sys.argv[0]), type='error')
		exit()
	if sys.argv[1].split('.')[-1] == 'docx':
		try:
			doc = docx.Document(sys.argv[1])
		except Exception:
			printLog('文件打开失败', type='error')
			exit()

		pending = set([i.text[i.text.find('http'):] for i in doc.paragraphs if i.text and 'bizconfstreaming' in i.text])

		while len(pending) != 0:
			removed = []
			for i in pending:
				printLog('正在检查{0}...'.format(i), back=True)
				try:
					if doWatchAndDownload(i):
						removed.append(i)
					time.sleep(1)
				except requests.exceptions.ConnectionError:
					time.sleep(60)
			for i in removed:
				pending.remove(i)
	else:
		while not doWatchAndDownload(sys.argv[1]):
			time.sleep(1)