import os
import pandas as pd
import requests
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
import string
from bs4 import BeautifulSoup

def createHtmlFiles( df, html_files_path ):
    os.makedirs( html_files_path )
    for index, row in df.iterrows():
        try:
            response = requests.get( row[ 'URL' ] )
            response.raise_for_status()
            cur_file = os.path.join( html_files_path, row[ 'URL_ID' ] + '.html' )
            with open( cur_file, 'a' ) as file:
                file.write( response.text )
        except requests.exceptions.RequestException as e:
            with open( 'Unavailable.txt', 'a' ) as file:
                file.write( row[ 'URL_ID' ] )
                file.write( '\n' )
    print( "HTML Files created")


def createTextFiles( text_files_path, html_files_path ):
    os.makedirs( text_files_path )
    for html_file in os.listdir( html_files_path ):
        text = ""
        html_file_path = os.path.join( html_files_path, html_file )
        with open( html_file_path, 'r', encoding='utf-8' ) as file:
            html_content = file.read()
            soup = BeautifulSoup( html_content, 'html.parser' )
            div_tag = soup.find( 'div', class_='tagdiv-type' )
            if div_tag:
                text += div_tag.get_text()
            else:
                text += "NOT FOUND"
        
        index = html_file.find( "." )
        text_file = html_file[ :index ]
        text_file_path = os.path.join( text_files_path, text_file + '.txt' )
        with open( text_file_path, 'a') as file:
            file.write( text )
    print( "Text Files created" )
    

def sentimentalAnalysis( all_words ):
    positive = []
    negative = []
    with open( './MasterDictionary/positive-words.txt', 'r', encoding='ISO-8859-1' ) as file:
        content = file.read()
        positive = word_tokenize( content )

    with open( './MasterDictionary/negative-words.txt', 'r', encoding='ISO-8859-1' ) as file:
        content = file.read()
        negative = word_tokenize( content )

    all_stop_words = []
    for stop_file in os.listdir( './StopWords' ):
        stop_file_path = os.path.join( './StopWords', stop_file )
        with open( stop_file_path, 'r', encoding='ISO-8859-1' ) as file:
            content = file.read()
            stop_words = word_tokenize( content )
            all_stop_words.extend( stop_words )

    words = [ word for word in all_words if word.upper() not in all_stop_words 
                and word not in string.punctuation ]

    positive_score = 0
    negative_score = 0

    for word in words:
        if word.lower() in positive:
            positive_score += 1
        if word.lower() in negative:
            negative_score += 1

    return positive_score, negative_score, len( words )

def main():
    df = pd.read_excel( './input.xlsx' )
    res = pd.DataFrame( columns=['URL_ID', 'URL', 'POSITIVE SCORE', 'NEGATIVE SCORE',
                                 'POLARITY SCORE', 'SUBJECTIVITY SCORE', 'AVG SENTENCE LENGTH',
                                 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX', 
                                 'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 
                                 'WORD COUNT', 'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 
                                 'AVG WORD LENGTH'] )

    html_files_path = os.path.expanduser( "./HtmlFiles/" )
    text_files_path = os.path.expanduser( "./TextFiles/" )
    if not os.path.exists( text_files_path ):
        if not os.path.exists( html_files_path ):
            createHtmlFiles( df, html_files_path )
        createTextFiles( text_files_path, html_files_path )

    unavailable_files = {}

    with open( 'Unavailable.txt', 'r' ) as file:
        content = file.read()
        unavailable_files = set( content.split() )

    for index, row in df.iterrows():
        res.loc[ index, 'URL_ID' ] = row[ 'URL_ID' ]
        res.loc[ index, 'URL' ] = row[ 'URL' ]
        if row[ 'URL_ID' ] in unavailable_files:
            res.loc[ index, 'POSITIVE SCORE' ] = "PAGE NOT FOUND"
            res.loc[ index, 'NEGATIVE SCORE' ] = "PAGE NOT FOUND"
            res.loc[ index, 'POLARITY SCORE' ] = "PAGE NOT FOUND"
            res.loc[ index, 'SUBJECTIVITY SCORE' ] = "PAGE NOT FOUND"
            res.loc[ index, 'AVG SENTENCE LENGTH' ] = "PAGE NOT FOUND"
            res.loc[ index, 'PERCENTAGE OF COMPLEX WORDS' ] = "PAGE NOT FOUND"
            res.loc[ index, 'FOG INDEX' ]  = "PAGE NOT FOUND"
            res.loc[ index, 'AVG NUMBER OF WORDS PER SENTENCE' ] = "PAGE NOT FOUND"
            res.loc[ index, 'COMPLEX WORD COUNT' ] = "PAGE NOT FOUND" 
            res.loc[ index, 'WORD COUNT' ] = "PAGE NOT FOUND"
            res.loc[ index, 'SYLLABLE PER WORD' ] = "PAGE NOT FOUND"
            res.loc[ index, 'PERSONAL PRONOUNS' ] = "PAGE NOT FOUND"
            res.loc[ index, 'AVG WORD LENGTH' ] = "PAGE NOT FOUND"
        else:
            with open( text_files_path + row[ 'URL_ID' ] + '.txt', 'r' ) as file:
                content = file.read()
                words = word_tokenize( content )
                sentences = sent_tokenize( content )

                # 1. Sentimental Analysis
                positive_score, negative_score, word_cnt = sentimentalAnalysis( words )
                res.loc[ index, 'POSITIVE SCORE' ] = positive_score
                res.loc[ index, 'NEGATIVE SCORE' ] = negative_score
                res.loc[ index, 'POLARITY SCORE' ] = ( ( positive_score - negative_score ) / 
                    ( positive_score + negative_score + 0.000001 ) )
                res.loc[ index, 'SUBJECTIVITY SCORE' ] = ( ( positive_score + 
                    negative_score ) / ( word_cnt + 0.000001 ) )

                # 3. Average Number of Words Per Sentence
                res.loc[ index, 'AVG NUMBER OF WORDS PER SENTENCE' ] = \
                    len( words ) / len( sentences )

                # 4. Complex Word Count, 5. Word Count, 6. Syllable Count Per Word
                # 7. Personal Pronouns, 8. Average Word Count
                complex_words = 0
                word_count = 0
                syllable_count = 0
                pronoun_count = 0
                letter_count = 0
                pronouns = { 'i', 'we', 'my', 'ours', 'us'}
                for word in words:
                    vowel_count = 0
                    for letter in word:
                        letter_count += 1
                        if letter in 'aeiou':
                            vowel_count += 1
                            syllable_count += 1
                    if vowel_count > 2:
                        complex_words += 1
                    if word[ -2: ] == "es" or word[ -2: ] == "ed":
                        syllable_count -= 1
                    if ( word not in string.punctuation and 
                         word not in stopwords.words( 'english' ) ):
                         word_count += 1
                    if word.lower() in pronouns:
                        pronoun_count += 1
                    if word == "US":
                        pronoun_count -= 1
                    
                res.loc[ index, 'COMPLEX WORD COUNT' ] = complex_words
                res.loc[ index, 'WORD COUNT'] = word_count
                res.loc[ index, 'SYLLABLE PER WORD' ] = syllable_count / len ( words )
                res.loc[ index, 'PERSONAL PRONOUNS'] = pronoun_count
                res.loc[ index, 'AVG WORD LENGTH' ] = letter_count / len( words )

                # 2. Analysis of Readability
                res.loc[ index, 'AVG SENTENCE LENGTH' ] = \
                    word_count / len( sentences )
                res.loc[ index, 'PERCENTAGE OF COMPLEX WORDS' ] = \
                    complex_words / word_count
                res.loc[ index, 'FOG INDEX' ] = 0.4 * ( res.loc[ index, 'AVG SENTENCE LENGTH' ]
                    + res.loc[ index, 'PERCENTAGE OF COMPLEX WORDS' ] )
                
    res.to_excel( 'Output.xlsx' )                      

if __name__ == "__main__":
    main()