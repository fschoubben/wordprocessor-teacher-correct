# wordprocessor-teacher-correct
Correct automaticaly as much as possible for the "technical" parts I ask them

# first alpha version 
Disclaimer : I didn't code for 20 years or so, so it's a ugly, but it works...

A lot of things still Todo, but from now, if you give it .pdf and .docx files in the same directory, it will check : 
* from the pdf, 
  * if filename correspond to a patter, 
  * if files < a size
...
  
And it produce an excel file- with
* 3 sheets, 2 groups and "the others"
* header with titles and maximum scores
* scores of each student, 
* a comment explaining why they don't have maximum points
* the total of each points
* a yellow background if I have to check something
* a footer with average, max and min points for each score