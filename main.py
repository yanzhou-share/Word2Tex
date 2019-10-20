import textract
import re

text = textract.process("Word.docx")

paragraph = text.decode()

# replace all EndNote citation to RN
citation = re.compile('[0-9a-zA-Z’]+,\s\w+\s#\w+')
# citation.findall(paragraph)
paragraph = citation.sub(lambda m: 'RN'+ m.group().split('#')[-1], paragraph)

# add all citation with \cite{}
citation_bracket = re.compile('\{\w+\;?\w+?\;?\w+?[0-9a-zA-Z;]+?\}')
#citation_bracket.findall(paragraph)
paragraph = citation_bracket.sub(lambda m: '\cite' + m.group().replace(';', ','), paragraph)

#print(paragraph)


math = re.compile('[0-9a-zA-Z;,\(\)]+_[0-9a-zA-Z;,\(\)]+\^?[0-9a-zA-Z;,\(\)\-]+?')
#print(math.findall(paragraph))
paragraph = math.sub(lambda m: '$' + m.group() + '$', paragraph)
#print(paragraph)


paragraph = paragraph.replace('SrTiO3', 'SrTiO$_3$').replace('LaAlO3', 'LaAlO$_3$')


with open("output.tex", "w") as file:
    file.write(paragraph)
