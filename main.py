import textract
import re

text = textract.process("Word.docx")

processed = text.decode()

# replace all EndNote citation to RN
citation = re.compile('[0-9a-zA-Z’]+,\s\w+\s#\w+')
# citation.findall(processed)
processed = citation.sub(lambda m: 'RN'+ m.group().split('#')[-1], processed)

# add all citation with \cite{}
citation_bracket = re.compile('\{\w+\;?\w+?\;?\w+?[0-9a-zA-Z;]+?\}')
#citation_bracket.findall(processed)
processed = citation_bracket.sub(lambda m: '\cite' + m.group().replace(';', ','), processed)

#print(processed)
math = re.compile('[0-9a-zA-Z;,\(\)]+_[0-9a-zA-Z;\(\)]+\^[0-9a-zA-Z;\(\)]+')
processed = math.sub(lambda m: '$' + m.group() + '$', processed)

math = re.compile('[0-9a-zA-Z;,\(\)]+_[0-9a-zA-Z;\(\)]+[^\^]')
processed = math.sub(lambda m: '$' + m.group() + '$', processed)

#print(processed)

processed = processed.replace('SrTiO3', 'SrTiO$_3$').replace('LaAlO3', 'LaAlO$_3$')
processed = processed.replace('CaTiO3', 'CaTiO$_3$')
processed = processed.replace('~', '$\sim$')

print(processed)

with open("output.tex", "w") as file:
    file.write(processed)
