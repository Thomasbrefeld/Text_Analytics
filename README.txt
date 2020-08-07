============== Text Analyics ==========================

Will generate an excel file with grouped topics and 
subtopics to improve productivity. This Application is 
made with Python and uses the NLTK wordnet library for 
synonyms. 

GitHub: https://github.com/Thomasbrefeld/Text_Analytics


-------------- Quick Notes ----------------------------

* A sentence found in multiple topics will be displayed
  in all relevant topics. Therefore the program will
  make duplicates of sentences.
* An input file can be .xlsx or .csv (UTF8)
* Input file assumes the first row as a header 
  (and is ignored)
* Program can handle words or phrases
* First column in Topic is the actual topic
* First column in the subtopic is the parent topic


-------------- Requirments ----------------------------

System: Windows 10
Programs: .xlsx reader (Microsoft Excel Preferred)


-------------- Exclude Words or Phrases ---------------

To exclude a sentence that contains a word or phrase
place a '-' dash or '/' forward slash in front of the
word or phrase to be removed. Note: you can only 
exclude words from the 'topic_categorize.csv' file
not 'subtopic_categorize.csv'

Example:
  To remove recal and sport from topic awd:
    AWD, four, 4x4, -sport, 4wd, utility, /recall


-------------- File Layouts ---------------------------

Input file format (*.xlsx, *.csv - UTF8):
  Header (Skips),
  sentance1,
  sentance2,
  sentance3,

Topic file format ('src/topic_categorize.csv'):
  Topic1, descriptor1, descriptor2, descriptor3
  Topic2, descriptor1, descriptor2, descriptor3
  Topic3, descriptor1, descriptor2, descriptor3

Subtopic file format ('src/subtopic_categorize.csv'):
  Topic1, Subtopic1, Subtopic2, Subtopic3
  Topic2, Subtopic1, Subtopic2, Subtopic3
  Topic3, Subtopic1, Subtopic2, Subtopic3


-------------- Error Codes ----------------------------

Code 2: Bad input file
  * Ensure the input file exsits and is of type .xlsx 
    or .csv with encoding of UTF8. 

Code 3: Bad output file
  * Ensure the output location is not open by another
    program and that you have permission to the
    selected file.
  * Check that output file is of type .xlsx.

Code 4: Bad topic file
  * Ensure that the topic file is of type .csv with
    encoding of UTF8.

Code 5: Bad subtopic file
  * Ensure that the topic file is of type .csv with
    encoding of UTF8.

Code 6: Failed reading Topic or subtopic 
  * Ensure that the topic file and subtopic have the
    encoding of UTF8 (defined when saving as .csv).
  * Check that the format follows as described above
    in 'File Layouts'.
  * Ensure there is no '\' backslash in either file
  * check that you have permission to read both files.

Code 7: Failed to sort input file
  * Ensure that the input file if .csv is encoded as
    UTF8 (defined when saving as .csv).
  * Ensure you have permission to access the file.
  * Check that the format follows as described above
    in 'File Layouts'.

Code 8: Unsupported File type
  * Ensure that the input, output, topic, and
    subtopic files exist and you have permission to
    access them.
  * Ensure there is no '\' backslash in any file.
  * Check the format of each file as described above
    in 'File Layouts'.

Code 9: Mismatched Topic file and Subtopic file
  * Follow error log in the main directory called
    'topic_error_log.txt'.
  * Ensure that Topic order is maintained between
    the Topic and Subtopic files.
  * Check spelling between of Topics between the
    Topic and Subtopic files.

Code 50: Uncaught Error in the main sort
  * Restart and run the program again.
  * Redownload the whole application.
  * If the issue persists email maintainer.


-------------- Contact --------------------------------

Please feel free to reach out if you have any
questions, concerns, or issues.

Name: Thomas Brefeld, Jr.
Email: thomas.brefeld@gmail.com