This is just a small script i done to cut some documentation work 

Basically i was required to key in datas available in a excel file into a word template and generate a "activity sheet"

Since i am learning python, i tried to let py do all the work instead of me copying and pasting hundred over times. 

Result : 
  this amazing script LOL . 
 
 -basically anyone can modify this script to help them key in excel file data into a available template file. 
 
 the more troublesome step is to update the para&run value ( python-docx's function, the para & run value will determind where you will add your data to on the doc file) 
 you can use some for loop function like:
 
	#add below if using py2.x
	from __future__ import print_function

	doc = docx.Document(‘Template.docx’）
	for i in range(len(doc.paragraphs)):
  		print('\n'+str(i)) #number display on top of the actual paragraph is the para value
		for j in range(len(doc.paragraphs[i].runs)):
			#number inside <> is the run value for the following words
			print('<' + str(j)+ '>' +doc.paragraphs[i].runs[j].text , end = '')


The best way to find the correct place to put your data value is edit the template document and add a placeholder (bold and italic it so it is a new run)

