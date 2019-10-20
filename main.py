#Simple script converts word.docx manuscript to latex. 

import textract
import re

text = textract.process("Word.docx")


paragraph = text.decode()

# replace all EndNote citation to RN
citation = re.compile('[0-9a-zA-Z’]+,\s\w+\s#\w+')
# citation.findall(paragraph)
paragraph_mid = citation.sub(lambda m: 'RN'+ m.group().split('#')[-1], paragraph)


citation_bracket = re.compile('\{\w+\;?\w+?\;?\w+?[0-9a-zA-Z;]+?\}')
citation_bracket.findall(paragraph_mid)
paragraph_end = citation_bracket.sub(lambda m: '\cite' + m.group().replace(';', ','), paragraph_mid)
print(paragraph_end)
