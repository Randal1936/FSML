import wordcloud
import os


def drawer(name, doc):
    w1 = wordcloud.WordCloud(font_path='simhei.ttf', width=1200, height=800, background_color='white')
    w1.generate(doc)
    w1.to_file(name + ".png")

