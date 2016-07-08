import sys
import subprocess
import fileinput
import string
import math

#"""###initialization"""

#"""displays various debugging messages along the way"""
debugging = False

#"""MaxDistanceFromEnd is the farthest the program is allowed to go from the middle point"""
#"""The actual max distance is 1 less than this variable. The same goes for MaxWindowExtension"""

DefaultSalt = 50

ExpectedNumberOfParameters = 6
MaxDistanceFromEnd = 30
MaxWindowExtension = 5

minOverlapTm = 48
primer_tm = 62
NumberOfOutputs = 1
MinimumWindow = 20
length0 = MinimumWindow

if debugging:
 print "# of args: " + str(len(sys.argv))
 print "script: " + sys.argv[0]
 print "input: " + sys.argv[1]
 print "RNAFold: " + sys.argv[2]
 print "TempDir: " + sys.argv[3]
 print


 
#"""full path to python script - not really important, for debugging"""
PythonScriptpath = sys.argv[0]

#"""full path to input file"""
InputFile = sys.argv[1]

#"""full executable path to RNAFold"""
RNAFoldPath = sys.argv[2]

TempFolder = 'C:/Temp/PythonTemp'
#"""JA: define the folder for the temp files"""
if len(sys.argv) > 3:
 if sys.argv[3] <> "":
  TempFolder = sys.argv[3]
 


TempFolder.replace('\\', '/')

OverlapListFile = TempFolder + '/folds_sr.tmp'
EnergiesFile = TempFolder + '/folds_dG.tmp'

if debugging:
 print TempFolder
 print OverlapListFile
 print EnergiesFile
 print

 
 
 
 
 
 
#"""Extensions"""
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
 
def ReverseComplement(s): 
 s = reverse(s) 
 s = complement(s) 
 return s 

#"""###read input file"""
seq1 = []
with open(InputFile, 'r') as TempFile:
 for line in TempFile:
  seq1.append(line[:-1])
 seq2 = seq1[0]+seq1[1]+seq1[2]
 a = len(seq1[0])+(len(seq1[1])/2)
 if len(seq1[1]) > MinimumWindow:
  window = MinimumWindow #this is the shit
  #if FlorianParameter <> '13'
  #window = len(seq1[1])
 else:
  window = MinimumWindow
 TempFile.close
 
 
 
 
 
 
 
 
 
#"""###Parse inputs"""
#"""MaxDistanceFromEnd is the farthest the program is allowed to go from the middle point"""
#"""The actual max distance is 1 less than this variable. The same goes for MaxWindowExtension"""
#"""TestFactor is a number of tested conditions that will be possible. If it's negative, the program tweaks the constraints a bit"""

TestFactor = len(seq2) - window - MaxDistanceFromEnd - MaxWindowExtension + 2

FlorianParameter = seq1[3]

if str(FlorianParameter) == '1':
 TestFactor = TestFactor - len(seq1[0])
elif str(FlorianParameter) == '2':
 if len(seq1[0])<=len(seq1[2]):
  TestFactor = TestFactor - len(seq1[0])
 else: 
  TestFactor = TestFactor - len(seq1[2])
elif str(FlorianParameter) == '3':
 TestFactor = TestFactor - len(seq1[2])
elif str(FlorianParameter) == '12': 
 TestFactor = TestFactor - len(seq1[0]+seq1[1])
elif str(FlorianParameter) == '23':
 TestFactor = TestFactor - len(seq1[1]+seq1[2])
elif str(FlorianParameter) == '13':
 TestFactor = TestFactor - len(seq1[0]+seq1[2])
else:
 TestFactor = TestFactor

if TestFactor <0:
 MaxDistanceFromEnd = MaxDistanceFromEnd + TestFactor

 

TemplateName1 = 'Template1'
TemplateName2 = 'Template2'

if len(seq1)>=ExpectedNumberOfParameters:
 if (str(seq1[4]) != "") and (str(seq1[5]) != "") and (str(seq1[4]) != str(seq1[5])):
  TemplateName1 = str(seq1[4])
  TemplateName2 = str(seq1[5])







  
  
#"""###construct inputs for RNAFold"""
seq3 = []

for j in range(0,MaxWindowExtension):
 length = length0 + j
 if str(FlorianParameter) == '1':
  allowed1 = 'Added Seq, Template2'
  for i in range(0,window):
   seq3.append(seq2[len(seq1[0])+i:len(seq1[0])+i+length]) 
 elif str(FlorianParameter) == '2':
  allowed1 = 'Template1, Template2'
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])+len(seq1[1])+i:len(seq1[0])+len(seq1[1])+i+length]) 
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])-i-length:len(seq1[0])-i])
 elif str(FlorianParameter) == '3':
  allowed1 = 'Template1, Added Seq'
  for i in range(0,window):
   seq3.append(seq2[len(seq1[0])+len(seq1[1])-i-length:len(seq1[0])+len(seq1[1])-i])
 elif str(FlorianParameter) == '12':
  allowed1 = 'Template2'
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])+len(seq1[1])+i:len(seq1[0])+len(seq1[1])+i+length]) 
 elif str(FlorianParameter) == '23':
  allowed1 = 'Template1'
  for i in range(0,MaxDistanceFromEnd):
   seq3.append(seq2[len(seq1[0])-i-length:len(seq1[0])-i])
 elif str(FlorianParameter) == '13':
  allowed1 = 'Added Seq'
  for i in range(0,window):
   seq3.append(seq2[len(seq1[0])+i:len(seq1[0])+i+length]) 
 else:
  allowed1 = 'Template1, Added Seq, Template2'
  for i in range(-window,window):
   seq3.append(seq2[a+i-(length/2):a+i-(length/2)+length])

seq4 = []
for item in seq3:
 seq4.append(ReverseComplement(item))


with open(OverlapListFile,'w') as TempFile:
 for i in range(0,len(seq3)):
  TempFile.write(seq3[i] + '\r\n')
  TempFile.write(seq4[i] + '\r\n') 
 TempFile.close
 
 
 
 
 
 
#"""###run RNAFold"""

#"""OS X Version"""
#"""command1 = subprocess.Popen(['RNAfold --noGU --noPS -T 10 < folds_sr.tmp > folds_dG.tmp'], shell=True)"""
#"""\OS X Version"""

#"""JA - Windows version"""
TempInput = file(OverlapListFile,'r')
TempOutput = file(EnergiesFile,'w')
#"""??TempInput"""
command1 = subprocess.Popen([RNAFoldPath, "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)
#"""command1 = subprocess.Popen(["RNAfold.exe", "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)"""
#"""command1 = subprocess.Popen(["RNAfold", "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)"""
#"""\Windows version"""

command1.wait()

TempInput.close()
TempOutput.close()






#"""###extract RNAFold outputs"""
L1 = []
with open(EnergiesFile,'r') as TempFile:
 i = 0
 for eachLine in TempFile:
  i += 1
  if i%2 == 0:
   L1.append(eachLine[-8:-2])
L2 = []
for i in range(0,len(L1)):
 if i%2 == 0:
  L2.append((seq3[i/2],float(L1[i])+float(L1[i+1])))
L3 = sorted(L2, key=lambda x: x[1], reverse=True)






#"""###output"""
   
  
  
#"""###output header"""
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







#"""###print Gibson overlaps"""

print 'Best overlapping segment(s):'

L4 = []
L5 = []
L6 = []
i = 0
j = 0
tempSeq = L3[0][0]
tempTm = Tm_dan(tempSeq,DefaultSalt)

if tempTm >= minOverlapTm:
 i = i + 1
 L4.append(tempSeq)
 L5.append(tempTm)
 L6.append(L3[0][1])
 
j = 0
while i < NumberOfOutputs:
 j = j + 1
 tempSeq = L3[j][0]
 tempTm = Tm_dan(tempSeq,DefaultSalt)
 if tempTm >= minOverlapTm:
  L4.append(tempSeq)
  L5.append(tempTm)
  L6.append(L3[j][1])
  i = i + 1
   
for i in range(0,NumberOfOutputs): 
  print L4[i] + '    dG: ' + str(L6[i]) + '    Tm: ' + str(L5[i])[:4]
print







 
#"""###calculate and print primer sequences"""
for i in range(0,NumberOfOutputs):
 if seq2.find(seq1[2]) >= seq2.find(L4[i]):
  bigger1 = seq2.find(seq1[2])
 else:
  bigger1 = seq2.find(L4[i])
 for j in range(0,MaxDistanceFromEnd):
  if (Tm_dan(seq2[bigger1:seq2.find(seq1[2])+16+j],DefaultSalt) >= primer_tm) and (seq2[seq2.find(seq1[2])+16+j-1].upper() == 'C' or seq2[seq2.find(seq1[2])+16+j-1].upper() =='G'):   
   break
  
 PrintTemp4 = str(Tm_dan(seq2[bigger1:seq2.find(seq1[2])+16+j],DefaultSalt))[:4]
 PrintTemp3 = seq2[seq2.find(L4[i]):seq2.find(seq1[2])+16+j]
 
 print '>Overlap ' + str(i+1) + ': ' + TemplateName2 + '_f  (Tm ' + PrintTemp4 + ')'
 print PrintTemp3
   
 seq2r = ReverseComplement(seq2)
 l3ir = ReverseComplement(L4[i])
 seq10r = ReverseComplement(seq1[0])
 
 if seq2r.find(seq10r) >= seq2r.find(l3ir):
  bigger1 = seq2r.find(seq10r)
 else:
  bigger1 = seq2r.find(l3ir)
  
 for j in range(0,MaxDistanceFromEnd):
  if (Tm_dan(seq2r[bigger1:seq2r.find(seq10r)+16+j],DefaultSalt) >= primer_tm) and (seq2r[seq2r.find(seq10r)+16+j-1].upper() == 'C' or seq2r[seq2r.find(seq10r)+16+j-1].upper() =='G'):      
   break
 print
 
 PrintTemp2 = str(Tm_dan(seq2r[bigger1:seq2r.find(seq10r)+16+j],DefaultSalt))[:4]
 PrintTemp1 = seq2r[seq2r.find(l3ir):seq2r.find(seq10r)+16+j]
 
 print '>Overlap ' + str(i+1) + ': ' + TemplateName1 +'_r  (Tm ' + PrintTemp2 + ')'
 print PrintTemp1
 
 
 
 
 
#"""###final tagged output that is read by Gibson Assembly excel macro"""  
print '[OVERLAP][OverlapSequence]' + L4[0] + '[\OverlapSequence] [dG]' + str(L6[0]) + '[\dG] [Tm]' + str(L5[0])[:4] + '[\Tm]'
print '[PRIMER1][PrimerName]' + TemplateName2 + '_f[\PrimerName] [Sequence]' + PrintTemp3 + '[\Sequence] [Tm]' + PrintTemp4 + '[\Tm]'
print '[PRIMER2][PrimerName]' + TemplateName1 + '_r[\PrimerName] [Sequence]' + PrintTemp1 + '[\Sequence] [Tm]' + PrintTemp2 + '[\Tm]'

