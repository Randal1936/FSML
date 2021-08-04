import numpy as np
import pandas as pd
import os


print('LDA Training now......')
tf_vectorizer = CountVectorizer(strip_accents='unicode',
                                stop_words='english',
                                lowercase=True,
                                token_pattern=r'\b[\u4e00-\u9fa5]{2,}\b',
                                max_df=0.5,
                                min_df=10)

dtm_tf = tf_vectorizer.fit_transform(words)

tfidf_vectorizer = TfidfVectorizer(**tf_vectorizer.get_params())
dtm_tfidf = tfidf_vectorizer.fit_transform(words)

lda_tf = LatentDirichletAllocation(n, random_state=0)
lda_tf.fit(dtm_tf)

# for TFIDF DTM
lda_tfidf = LatentDirichletAllocation(n, random_state=0)
lda_tfidf.fit(dtm_tfidf)

d = pyLDAvis.sklearn.prepare(lda_tf, dtm_tf, tf_vectorizer)

pyLDAvis.save_html(d, '银保监_lda_'+str(n)+'.html')







