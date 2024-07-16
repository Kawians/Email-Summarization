# Email-Summarization
The main purpose of this project is to make understanding of back and forth emails easier for employees who are dealing with a numerous number of emails during the day.

### What is the basic challenge?
The challenge which arose my interst to work on this matter was to getting to know that on some email platforms, conversations are not attached to each other as an email chain. So, as a result, if two person have had a 20 back and forth email, it would be resulted in 20 separated email files which are hard to review.

Also, what if another person is going to get involved in that chain? It would be much more difficult for them to understand what has happened!!

### My solution to this
What I have done was not that complicated. The script I wrote goes to the folder which all of the back and forth emails are stored there and do the following two things:

1- It combines all of the emails in one word doc, in order for everyone to be able to go over all of them easily.
<br>2- Most important, it generates another doc file which includes a summerization of all of the emails on Archive folder. Then anyone who is reaching out to the conversation could be able to understand what is happennig in just one paragraph.

### Technical Aspects
In order to do this project, the following libraries are implemented:
<br>
<br>1- nltk for the NLP process and summarization.
<br>2- email library to read and parsing the eml (outlook format of email files).
<br>3- docx in order for generating .docx outputs.

### How to use?
It's very simple. You only need to pip install email, nltk, and docx.
<br>Then you need to change the following three lines in the code and put your own local path:
<br>folder_path = 'Path to Email Archive folder'
<br>output_combined = 'path to output combined_emails.docx'
<br>output_summary = 'path to output summary_emails.docx'
<br>
<br>and running the script will provide you with the result.

### Result:
The result could be seen by openning combined_emails.docx and summary_emails.docx.