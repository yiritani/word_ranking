import re
from collections import Counter

import MeCab
import openpyxl
from matplotlib import pyplot as plt
from wordcloud import WordCloud

# 取り出したい品詞
select_conditions = ['動詞', '形容詞', '名詞']
tagger = MeCab.Tagger('')
tagger.parse('')
hiragana_re = re.compile('[\u3041-\u309F]+')

input_file_path = 'input/text.txt'

# 画像設定
font_path = "/Library/Fonts/Arial Unicode.ttf"
wc_img_file_name = "media/img/wc_image_en.png"
background_color = "white"

# エクセル設定
excel_filename = 'word_ranking.xlsx'
sheet_name = 'result'


def wakati_text(input_text):
	node = tagger.parseToNode(input_text)
	terms = []

	while node:
		# 単語
		term = node.surface
		# 品詞
		pos = node.feature.split(',')[0]
		if pos in select_conditions:
			if hiragana_re.match(term) and len(term) <= 1:
				node = node.next
				continue
			terms.append(term)
		node = node.next

	text_result = ' '.join(terms)
	return text_result, terms


def generate_wordcloud(cloud_text):
	wc = WordCloud(background_color=background_color, font_path=font_path)
	wc.generate(cloud_text)
	plt.imshow(wc)
	wc.to_file(wc_img_file_name)


def generate_excel(input_text):
	# 二次元配列をエクセルに書き込む
	def write_list_2d(sheet, l_2d, start_row, start_col):
		for y, row in enumerate(l_2d):
			for x, cell in enumerate(row):
				sheet.cell(row=start_row + y,
						   column=start_col + x,
						   value=l_2d[y][x])

	wb = openpyxl.Workbook()
	sheet = wb.active
	sheet.title = sheet_name
	write_list_2d(sheet, input_text, 1, 1)

	img_to_excel = openpyxl.drawing.image.Image(wc_img_file_name)
	sheet.add_image(img_to_excel, 'D3')
	wb.save(excel_filename)


with open(input_file_path, 'r') as f:
	text = f.read()

joined_words, listed_words = wakati_text(text)
listed_words.sort()
counter_words = [list(i) for i in Counter(listed_words).most_common()]
generate_wordcloud(joined_words)
generate_excel(counter_words)
