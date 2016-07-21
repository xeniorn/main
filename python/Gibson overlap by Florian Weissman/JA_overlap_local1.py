import sys
import subprocess
import fileinput
import string
import math


"""MaxDistanceFromEnd is the farthest the program is allowed to go from the middle point"""
"""The actual max distance is 1 less than this variable. The same goes for MaxWindowExtension"""

ExpectedNumberOfParameters = 6
MaxDistanceFromEnd = 30
MaxWindowExtension = 1

primer_tm = 62
NumberOfOutputs = 1
MinimumWindow = 20
length0 = MinimumWindow


def Tm_dan(s,saltc,dnac=500):
 """Returns DNA mt using nearest neighbor thermodynamics. dnac is
 DNA concentration [nM] and saltc is salt concentration [mM]
 Works as EMBOSS dan, taken from (Breslauer et al. Proc. Natl. Acad.
 Sci. USA 83, 3746-3750 and Baldino et al. Methods in Enzymol. 168, 761-777). 
 Overcount function thanks to Greg Singer <singerg@xxxxxx>, rest by
 Sebastian Bassi <sbassi@xxxxxxxxxxxxxxxxxx>"""

 def overcount(st,p):
  """Returns how many p are on st, works even for overlapping"""
  ocu=0
  x=0
  while 1:
   try:
    i=st.index(p,x)
   except ValueError:
    break
   ocu=ocu+1
   x=i+1
  return ocu

 sup=string.upper(s)
 r = 1.98717
 LogDNA = r * math.log(dnac/4e9)
 vh=0
 vs=0

 vh=(overcount(sup,"AA"))*7.9+(overcount(sup,"TT"))*7.9+(overcount(sup,"AT"))*7.2+(overcount(sup,"TA"))*7.2+(overcount(sup,"CA"))*8.5+(overcount(sup,"TG"))*8.5+(overcount(sup,"GT"))*8.4+(overcount(sup,"AC"))*8.4

 vh=vh+(overcount(sup,"CT"))*7.8+(overcount(sup,"AG"))*7.8+(overcount(sup,"GA"))*8.2+(overcount(sup,"TC"))*8.2

 vh=vh+(overcount(sup,"CG"))*10.6+(overcount(sup,"GC"))*9.8+(overcount(sup,"GG"))*8+(overcount(sup,"CC"))*8

 vs=(overcount(sup,"AA"))*22.2+(overcount(sup,"TT"))*22.2+(overcount(sup,"AT"))*20.4+(overcount(sup,"TA"))*21.3

 vs=vs+(overcount(sup,"CA"))*22.7+(overcount(sup,"TG"))*22.7+(overcount(sup,"GT"))*22.4+(overcount(sup,"AC"))*22.4

 vs=vs+(overcount(sup,"CT"))*21.0+(overcount(sup,"AG"))*21.0+(overcount(sup,"GA"))*22.2+(overcount(sup,"TC"))*22.2
 vs=vs+(overcount(sup,"CG"))*27.2+(overcount(sup,"GC"))*24.4+(overcount(sup,"GG"))*19.9+(overcount(sup,"CC"))*19.9

 entropy = -10.8 - vs
 entropy = entropy + ((len(s)-1) * (math.log10(saltc/1000.0))*0.368)
 dTm = ((-vh*1000) / (entropy+LogDNA)) - 273.15
# print vh,vs,entropy
 return dTm
def reverse(s): 
 letters = list(s) 
 letters.reverse() 
 return ''.join(letters) 
def complement(s): 
 """OS X Version"""
 """basecomplement = {'A': 'T', 'C': 'G', 'G': 'C', 'T': 'A', 'U': 'A', 'a': 't', 'c': 'g', 'g': 'c', 't': 'a', 'u': 'a', '\n': '', ';': ''}  linux/mac"""
 """\OS X Version"""
 
 """JA - Windows version, adds \r to endline"""
 basecomplement = {'A': 'T', 'C': 'G', 'G': 'C', 'T': 'A', 'U': 'A', 'a': 't', 'c': 'g', 'g': 'c', 't': 'a', 'u': 'a', '\n': '', '\r': '', ';': ''}
 """\Windows version"""
 letters = list(s) 
 letters = [basecomplement[base] for base in letters] 
 return ''.join(letters) 
def gc(s): 
 gc = s.count('G') + s.count('C') 
 return gc * 100.0 / len(s) 
def reversecomplement(s): 
 s = reverse(s) 
 s = complement(s) 
 return s 


seq1 = []
for line in fileinput.input():
 seq1.append(line[:-1])
seq2 = seq1[0]+seq1[1]+seq1[2]
a = len(seq1[0])+(len(seq1[1])/2)
if len(seq1[1]) > MinimumWindow:
 window = len(seq1[1])
else:
 window = MinimumWindow
 
"""JA addition"""
"""MaxDistanceFromEnd is the farthest the program is allowed to go from the middle point"""
"""The actual max distance is 1 less than this variable. The same goes for MaxWindowExtension"""

TestFactor = len(seq2) - window - MaxDistanceFromEnd - MaxWindowExtension + 2

if str(seq1[3]) == '1':
 TestFactor = TestFactor - len(seq1[0])
elif str(seq1[3]) == '2':
 if len(seq1[0])<=len(seq1[2]):
  TestFactor = TestFactor - len(seq1[0])
 else:	
  TestFactor = TestFactor - len(seq1[2])
elif str(seq1[3]) == '3':
 TestFactor = TestFactor - len(seq1[2])
elif str(seq1[3]) == '12': 
 TestFactor = TestFactor - len(seq1[0]+seq1[1])
elif str(seq1[3]) == '23':
 TestFactor = TestFactor - len(seq1[1]+seq1[2])

if TestFactor <0:
 MaxDistanceFromEnd = MaxDistanceFromEnd + TestFactor	
	

TemplateName1 = 'Template1'
TemplateName2 = 'Template2'

if len(seq1)>=ExpectedNumberOfParameters:
	if (str(seq1[4]) != "") and (str(seq1[5]) != "") and (str(seq1[4]) != str(seq1[5])):
		TemplateName1 = str(seq1[4])
		TemplateName2 = str(seq1[5])


"""\JA addition"""

seq3 = []
if str(seq1[3]) == '1':
 allowed1 = 'Added Seq, Template2'
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(0,window):
   seq3.append(seq2[len(seq1[0])+i:len(seq1[0])+i+length]) 
elif str(seq1[3]) == '2':
 allowed1 = 'Template1, Template2'
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])+len(seq1[1])+i:len(seq1[0])+len(seq1[1])+i+length]) 
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])-i-length:len(seq1[0])-i])
elif str(seq1[3]) == '3':
 allowed1 = 'Template1, Added Seq'
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(0,window):
   seq3.append(seq2[len(seq1[0])+len(seq1[1])-i-length:len(seq1[0])+len(seq1[1])-i])
elif str(seq1[3]) == '12':
 allowed1 = 'Template2'
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])+len(seq1[1])+i:len(seq1[0])+len(seq1[1])+i+length]) 
elif str(seq1[3]) == '23':
 allowed1 = 'Template1'
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])-i-length:len(seq1[0])-i])
else:
 allowed1 = 'Template1, Added Seq, Template2'
 for j in range(0,MaxWindowExtension):
  length = length0 + j
  for i in range(-window,window):
   seq3.append(seq2[a+i-(length/2):a+i-(length/2)+length])


seq4 = []
for item in seq3:
 seq4.append(reversecomplement(item))


with open('folds_sr.tmp','w') as f:
 for i in range(0,len(seq3)):
  f.write(seq3[i] + '\r\n')
  f.write(seq4[i] + '\r\n')

"""OS X Version"""
"""command1 = subprocess.Popen(['RNAfold --noGU --noPS -T 10 < folds_sr.tmp > folds_dG.tmp'], shell=True)"""
"""\OS X Version"""

"""JA - Windows version"""
TempInput = file("folds_sr.tmp","r")
TempOutput = file("folds_dG.tmp","w")
command1 = subprocess.Popen(["RNAfold_v2.1.9.exe", "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)
"""command1 = subprocess.Popen(["RNAfold.exe", "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)"""
"""command1 = subprocess.Popen(["RNAfold", "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)"""
"""\Windows version"""

command1.wait()

TempInput.close()
TempOutput.close()

l1 = []
with open('folds_dG.tmp','r') as f:
 i = 0
 for eachLine in f:
  i += 1
  if i%2 == 0:
   l1.append(eachLine[-8:-2])
l2 = []
for i in range(0,len(l1)):
 if i%2 == 0:
  l2.append((seq3[i/2],float(l1[i])+float(l1[i+1])))
l3 = sorted(l2, key=lambda x: x[1], reverse=True)



print;print;print


print 'Overlap allowed in: ' + str(allowed1)
print
print 'Input sequences:'
print 'Template1:         ' + seq1[0]
print 'Built by oligos:   ' + seq1[1]
print 'Template2:         ' + seq1[2]
#for item in seq1:
# print item
print
print 'Product:'
print seq2
print

'''
for item in seq3:
 print item
print
'''
print 'Best overlapping segment(s):'
for i in range(0,NumberOfOutputs):
 print l3[i][0] + '    dG: ' + str(l3[i][1]) + '    Tm: ' + str(Tm_dan(l3[i][0],50))[:4]
print
print 'Confirm with Zipfold:'
for i in range(0,NumberOfOutputs):
 print l3[i][0] + ';'
 print reversecomplement(l3[i][0]) + ';'
print

for i in range(0,NumberOfOutputs):
 if seq2.find(seq1[2]) >= seq2.find(l3[i][0]):
  bigger1 = seq2.find(seq1[2])
 else:
  bigger1 = seq2.find(l3[i][0])
 for j in range(0,MaxDistanceFromEnd):
  if (Tm_dan(seq2[bigger1:seq2.find(seq1[2])+16+j],50) >= primer_tm) and (seq2[seq2.find(seq1[2])+16+j-1].upper() == 'C' or seq2[seq2.find(seq1[2])+16+j-1].upper() =='G'):
   PrintTemp4 = str(Tm_dan(seq2[bigger1:seq2.find(seq1[2])+16+j],50))[:4]
   print '>Overlap ' + str(i+1) + ': ' + TemplateName2 + '_for  (Tm ' + PrintTemp4 + ')'
   PrintTemp3 = seq2[seq2.find(l3[i][0]):seq2.find(seq1[2])+16+j]
   print PrintTemp3
   break
 seq2r = reversecomplement(seq2)
 l3ir = reversecomplement(l3[i][0])
 seq10r = reversecomplement(seq1[0])
 if seq2r.find(seq10r) >= seq2r.find(l3ir):
  bigger1 = seq2r.find(seq10r)
 else:
  bigger1 = seq2r.find(l3ir)
 for j in range(0,MaxDistanceFromEnd):
  if (Tm_dan(seq2r[bigger1:seq2r.find(seq10r)+16+j],50) >= primer_tm) and (seq2r[seq2r.find(seq10r)+16+j-1].upper() == 'C' or seq2r[seq2r.find(seq10r)+16+j-1].upper() =='G'):
   PrintTemp2 = str(Tm_dan(seq2r[bigger1:seq2r.find(seq10r)+16+j],50))[:4]
   print '>Overlap ' + str(i+1) + ': ' + TemplateName1 +'_rev  (Tm ' + PrintTemp2 + ')'
   PrintTemp1 = seq2r[seq2r.find(l3ir):seq2r.find(seq10r)+16+j]
   print PrintTemp1
   break
 print

print '[OVERLAP][OverlapSequence]' + l3[0][0] + '[\OverlapSequence] [dG]' + str(l3[0][1]) + '[\dG] [Tm]' + str(Tm_dan(l3[0][0],50))[:4] + '[\Tm]'
print '[PRIMER1][PrimerName]' + TemplateName2 + '_for[\PrimerName] [Sequence]' + PrintTemp3 + '[\Sequence] [Tm]' + PrintTemp4 + '[\Tm]'
print '[PRIMER2][PrimerName]' + TemplateName1 + '_rev[\PrimerName] [Sequence]' + PrintTemp1 + '[\Sequence] [Tm]' + PrintTemp2 + '[\Tm]'
 
