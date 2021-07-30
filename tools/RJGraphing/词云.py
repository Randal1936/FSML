import wordcloud
import os


def drawer(name, df):
    m = []
    for i in range(0, len(df)):
        m.append(df.iloc[i])
    f = ' '.join(m)
    w1 = wordcloud.WordCloud(font_path='simhei.ttf', width=1200, height=800, background_color='white')
    w1.generate(f)
    w1.to_file(name + ".png")


drawer('10+次', df.loc[df[df['Sum'] > 10].index]['正文'])

