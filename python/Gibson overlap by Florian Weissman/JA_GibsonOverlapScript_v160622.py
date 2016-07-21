import sys
import subprocess
import fileinput
import string
import math
 
#"""Extensions"""
def DNA_Tm(s,saltc,dnac=500):
 """Returns DNA mt using nearest neighbor thermodynamics. dnac is
 DNA concentration [nM] and saltc is salt concentration [mM]
 Works as EMBOSS dan, taken from (Breslauer et al. Proc. Natl. Acad.
 Sci. USA 83, 3746-3750 and Baldino et al. Methods in Enzymol. 168, 761-777). 
 Overcount function thanks to Greg Singer <singerg@xxxxxx>, rest by
 Sebastian Bassi <sbassi@xxxxxxxxxxxxxxxxxx>"""
 
 def overcount(st,p):
  """Returns how many p are on st, works even for overlapping"""
  ocu = 0
  x = 0
  while 1:
   try:
    i = st.index(p,x)
   except ValueError:
    break
   ocu = ocu + 1
   x = i + 1
  return ocu
 
 
 sup = string.upper(s)
 r = 1.98717
 LogDNA = r * math.log(dnac / 4e9)
 vh = 0
 vs = 0
 
 vh = (overcount(sup,"AA")) * 7.9 + (overcount(sup,"TT")) * 7.9 + (overcount(sup,"AT")) * 7.2 + (overcount(sup,"TA")) * 7.2 \
+(overcount(sup,"CA")) * 8.5 + (overcount(sup,"TG")) * 8.5 + (overcount(sup,"GT")) * 8.4 + (overcount(sup,"AC")) * 8.4 \
+(overcount(sup,"CT")) * 7.8 + (overcount(sup,"AG")) * 7.8 + (overcount(sup,"GA")) * 8.2 + (overcount(sup,"TC")) * 8.2 \
+(overcount(sup,"CG")) * 10.6 + (overcount(sup,"GC")) * 9.8 + (overcount(sup,"GG")) * 8 + (overcount(sup,"CC")) * 8 \
 
 vs = (overcount(sup,"AA")) * 22.2 + (overcount(sup,"TT")) * 22.2 + (overcount(sup,"AT")) * 20.4 + (overcount(sup,"TA")) * 21.3 \
 +(overcount(sup,"CA")) * 22.7 + (overcount(sup,"TG")) * 22.7 + (overcount(sup,"GT")) * 22.4 + (overcount(sup,"AC")) * 22.4 \
 +(overcount(sup,"CT")) * 21.0 + (overcount(sup,"AG")) * 21.0 + (overcount(sup,"GA")) * 22.2 + (overcount(sup,"TC")) * 22.2 \
 +(overcount(sup,"CG")) * 27.2 + (overcount(sup,"GC")) * 24.4 + (overcount(sup,"GG")) * 19.9 + (overcount(sup,"CC")) * 19.9 \
 
 entropy = -10.8 - vs
 entropy = entropy + ((len(s) - 1) * (math.log10(saltc / 1000.0)) * 0.368)
 dTm = ((-vh * 1000) / (entropy + LogDNA)) - 273.15
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
 basecomplement = {'A': 'T', 'C': 'G', 'G': 'C', 'T': 'A', 'U': 'A', 'a': 't', 'c': 'g', 'g': 'c', 't': 'a', 'u': 'a', '\n': '', '\r': '', ';': '',' ' : ''}
 """\Windows version"""

 result = ''

 if len(s) > 0:
  letters = list(s) 
  letters = [basecomplement[base] for base in letters] 
  result = result.join(letters)  

 return result
 
def gc(s): 
 gc = s.count('G') + s.count('C') 
 return gc * 100.0 / len(s) 
 
def ReverseComplement(s): 
 s = reverse(s) 
 s = complement(s) 
 return s 





#################################################################################################
#################################################################################################
#################################################################################################
#################################################################################################
#################################################################################################
#################################################################################################

#"""###initialization"""

#"""displays various debugging messages along the way"""
debugging = False

#"""MaxDistanceFromEnd is the farthest the program is allowed to go from the
#middle point"""
#"""The actual max distance is 1 less than this variable.  The same goes for
#MaxWindowExtension"""
DefaultSalt = 50

ExpectedNumberOfParameters = 6

NumberOfOutputs = 1

MinimumVariantsNumber = 5

MinimumGibsonOverlap = 20
MaxGibsonOverlap = 25
minGibsonOverlapTm = 50

MaxGibsonOverlapDistance = 35


MinPrimerAnneal = 15
minPrimerTm = 60

MaxPrimerAnneal = 30


MaxGibsonExtension = MaxGibsonOverlap - MinimumGibsonOverlap
MaxPrimerExtension = MaxPrimerAnneal - MinPrimerAnneal


if debugging:
 print "# of args: " + str(len(sys.argv))
 print "script: " + sys.argv[0]
 print "input: " + sys.argv[1]
 print "RNAFold: " + sys.argv[2]
 print "TempDir: " + sys.argv[3]
 print


 
#"""full path to python script - not really important, for debugging"""
PythonScriptPath = sys.argv[0]

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

 
 
 
 





#"""###read input file"""
InputLinesList = []
with open(InputFile, 'r') as TempFile:
 for line in TempFile:
  InputLinesList.append(line[:-1])
 TempFile.close

FirstSequence = InputLinesList[0]
InsertSequence = InputLinesList[1]
LastSequence = InputLinesList[2]
FlorianParameter = InputLinesList[3]

if len(InputLinesList) >= ExpectedNumberOfParameters:
 if (str(InputLinesList[4]) != "") and (str(InputLinesList[5]) != ""):
  if (str(InputLinesList[4]) != str(InputLinesList[5])):
   TemplateName1 = str(InputLinesList[4])
   TemplateName2 = str(InputLinesList[5])
  else:
   TemplateName1 = str(InputLinesList[4]) + "_1"
   TemplateName2 = str(InputLinesList[4]) + "_2"
 else:
  TemplateName1 = 'Template1'
  TemplateName2 = 'Template2'
else:
 """hi, hihihi, I do nothing!"""


FinalSequence = FirstSequence + InsertSequence + LastSequence
IndexOfMiddle = len(FirstSequence) + (len(InsertSequence) // 2)

#if len(InsertSequence) > MinimumGibsonOverlap:
# window = MinimumGibsonOverlap #this is the shit
#                         #if FlorianParameter <> '13'
#                         #window = len(InsertSequence)
#else:
window = MinimumGibsonOverlap
 
 

 
#"""###Parse inputs"""
#"""MaxDistanceFromEnd is the farthest the program is allowed to go from the
#middle point"""
#"""The actual max distance is 1 less than this variable.  The same goes for
#MaxWindowExtension"""
#"""TestFactor is a number of tested conditions that will be possible.  If it's
#negative, the program tweaks the constraints a bit"""
TestFactor = len(FinalSequence) - window - MaxGibsonOverlapDistance - MaxGibsonExtension + 2


if str(FlorianParameter) == '1':
 TestFactor = TestFactor - len(FirstSequence)
elif str(FlorianParameter) == '2':
 if len(FirstSequence) <= len(LastSequence):
  TestFactor = TestFactor - len(FirstSequence)
 else: 
  TestFactor = TestFactor - len(LastSequence)
elif str(FlorianParameter) == '3':
 TestFactor = TestFactor - len(LastSequence)
elif str(FlorianParameter) == '12': 
 TestFactor = TestFactor - len(FirstSequence + InsertSequence)
elif str(FlorianParameter) == '23':
 TestFactor = TestFactor - len(InsertSequence + LastSequence)
elif str(FlorianParameter) == '13':
 TestFactor = TestFactor - len(FirstSequence + LastSequence)
else:
 TestFactor = TestFactor

if TestFactor < 0:
 MaxGibsonOverlapDistance = MaxGibsonOverlapDistance + TestFactor



  
  
#"""###construct inputs for RNAFold"""
tempOverlapsList = []
indexlist = []
startindex = 0

for j in range(0,MaxGibsonExtension):
 length = MinimumGibsonOverlap + j
 if str(FlorianParameter) == '1':
  TextAllowed = 'Added Seq, Template2'
  for i in range(0,MaxGibsonOverlapDistance):
   startindex = len(FirstSequence) + i
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex)

 elif str(FlorianParameter) == '2':
  TextAllowed = 'Template1, Template2'
  tempMaxDistance = MaxGibsonOverlapDistance - len(InsertSequence)
  if tempMaxDistance < MinimumVariantsNumber:
   tempMaxDistance = MinimumVariantsNumber
  for i in range(0,tempMaxDistance):
   startindex = len(FirstSequence) + len(InsertSequence) + i
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex) 
  for i in range(0,tempMaxDistance):
   startindex = len(FirstSequence) - i - length
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex)

 elif str(FlorianParameter) == '3':
  TextAllowed = 'Template1, Added Seq'
  for i in range(0,MaxGibsonOverlapDistance):
   startindex = len(FirstSequence) + len(InsertSequence) - i - length
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex)

 elif str(FlorianParameter) == '12':
  TextAllowed = 'Template2'
  tempMaxDistance = MaxGibsonOverlapDistance - len(InsertSequence)
  if tempMaxDistance < MinimumVariantsNumber:
   tempMaxDistance = MinimumVariantsNumber
  for i in range(0,tempMaxDistance):
   startindex = len(FirstSequence) + len(InsertSequence) + i
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex)

 elif str(FlorianParameter) == '23':
  TextAllowed = 'Template1'
  tempMaxDistance = MaxGibsonOverlapDistance - len(InsertSequence)
  if tempMaxDistance < MinimumVariantsNumber:
   tempMaxDistance = MinimumVariantsNumber  
  for i in range(0,tempMaxDistance):
   startindex = len(FirstSequence) - i - length
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex)

 elif str(FlorianParameter) == '13':
  TextAllowed = 'Added Seq'
  tempMaxDistance = (len(InsertSequence) - length) // 2
  tempAdd = (len(InsertSequence) - length) % 2
  for i in range(-tempMaxDistance,tempMaxDistance + tempAdd):
   startindex = IndexOfMiddle
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex) 

 else:
  TextAllowed = 'Template1, Added Seq, Template2'
  tempMaxDistance = MaxGibsonOverlapDistance // 2
  for i in range(-tempMaxDistance , tempMaxDistance +1):
   startindex = IndexOfMiddle + i - length
   tempOverlapsList.append(FinalSequence[startindex:startindex + length]) 
   indexlist.append(startindex)

tempRCOverlapsList = []
for item in tempOverlapsList:
 tempRCOverlapsList.append(ReverseComplement(item))


with open(OverlapListFile,'w') as TempFile:
 for i in range(0,len(tempOverlapsList)):
  TempFile.write(tempOverlapsList[i] + '\r\n')
  TempFile.write(tempRCOverlapsList[i] + '\r\n') 
 TempFile.close
 
 
 
 
 
 
#"""###run RNAFold"""

#"""OS X Version"""
#"""command1 = subprocess.Popen(['RNAfold --noGU --noPS -T 10 < folds_sr.tmp >
#folds_dG.tmp'], shell=True)"""
#"""\OS X Version"""

#"""JA - Windows version"""
TempInput = file(OverlapListFile,'r')
TempOutput = file(EnergiesFile,'w')
#"""??TempInput"""
command1 = subprocess.Popen([RNAFoldPath, "--noGU", "--noPS", "-T 10"], stdin=TempInput, stdout=TempOutput)
#"""command1 = subprocess.Popen(["RNAfold.exe", "--noGU", "--noPS", "-T 10"],
#stdin=TempInput, stdout=TempOutput)"""
#"""command1 = subprocess.Popen(["RNAfold", "--noGU", "--noPS", "-T 10"],
#stdin=TempInput, stdout=TempOutput)"""
#"""\Windows version"""
command1.wait()

TempInput.close()
TempOutput.close()






#"""###extract RNAFold outputs"""
EnergyList = []
with open(EnergiesFile,'r') as TempFile:
 i = 0
 for eachLine in TempFile:
  i += 1
  if i % 2 == 0:
   EnergyList.append(eachLine[-8:-2])

ParsedEnergyList = []
for i in range(0,len(EnergyList)):
 if i % 2 == 0:
  ParsedEnergyList.append((tempOverlapsList[i / 2],float(EnergyList[i]) + float(EnergyList[i + 1]),indexlist[i / 2]))

SortedEnergyList = sorted(ParsedEnergyList, key=lambda x: x[1], reverse=True)






#"""###output"""
   
  
  
#"""###output header"""
print
print
print

print 'Overlap allowed in: ' + str(TextAllowed)
print
print 'Input sequences:'
print 'Template1:         ' + FirstSequence
print 'Built by oligos:   ' + InsertSequence
print 'Template2:         ' + LastSequence
#for item in InputLinesList:
# print item
print
print 'Product:'
print FinalSequence
print







#"""###print Gibson overlaps"""
print 'Best overlapping segment(s):'

OverlapsList = []
TmList = []
FinalEnergyList = []
OverlapIndexList = []
i = 0
j = 0
tempSeq = SortedEnergyList[0][0]
tempTm = DNA_Tm(tempSeq,DefaultSalt)

if tempTm >= minGibsonOverlapTm:
 i = i + 1
 OverlapsList.append(tempSeq)
 TmList.append(tempTm)
 FinalEnergyList.append(SortedEnergyList[0][1])
 OverlapIndexList.append(SortedEnergyList[j][2])
 
j = 0
while i < NumberOfOutputs:
 j = j + 1
 tempSeq = SortedEnergyList[j][0]
 tempTm = DNA_Tm(tempSeq,DefaultSalt)
 if tempTm >= minGibsonOverlapTm:
  OverlapsList.append(tempSeq)
  TmList.append(tempTm)
  FinalEnergyList.append(SortedEnergyList[j][1])
  OverlapIndexList.append(SortedEnergyList[j][2])
  i = i + 1
   
for i in range(0,NumberOfOutputs): 
  print OverlapsList[i] + '    dG: ' + str(FinalEnergyList[i]) + '    Tm: ' + str(TmList[i])[:4]
print







LastFragmentIndex = len(FirstSequence) + len(InsertSequence)
RCFirstFragmentIndex = len(LastSequence) + len(InsertSequence)

#"""###calculate and print primer sequences"""
for i in range(0,NumberOfOutputs):
 #LEGACY CODE FROM FLORIAN: fails if there are repeat sequences
 """
 if IndexOfFragTwo >= seq2.find(L4[i]):
  bigger1 = IndexOfFragTwo
 else:
  bigger1 = seq2.find(L4[i])
 """
 #new code JA
 #location of overlap in the sequence
 #if the overlap is inside fragment 2, then start counting the overlap residues
 #from the overlap (PCR will go from there, truncating it!)
 #else, start counting from the beginning of the fragment 2, where the primer
 #will actually anneal
 
 PrimerStartIndex = OverlapIndexList[i]

 if LastFragmentIndex > PrimerStartIndex:
  PrimerAnnealStartIndex = LastFragmentIndex
 else:
  PrimerAnnealStartIndex = PrimerStartIndex
  

 for tPrimerLength in range(MinPrimerAnneal,MaxPrimerAnneal + 1):
  FragmentEndIndex = PrimerAnnealStartIndex + tPrimerLength
  TerminalResidue = FinalSequence[FragmentEndIndex].upper()
  tAnnealSequence = FinalSequence[PrimerAnnealStartIndex:FragmentEndIndex + 1]
  tempTm = DNA_Tm(tAnnealSequence,DefaultSalt)
  if (tempTm >= minPrimerTm) and (TerminalResidue == 'C' or TerminalResidue == 'G'):   
   break
  
 PrintTm_for = str(tempTm)[:4]
 PrintSeq_for = FinalSequence[PrimerStartIndex:FragmentEndIndex + 1]
 
 print '>Overlap ' + str(i + 1) + ': ' + TemplateName2 + '_f  (Tm ' + PrintTm_for + ')'
 print PrintSeq_for
 
#repeat the same for reverse!
  
 RCFinalSequence = ReverseComplement(FinalSequence)
 RCtempOverlap = ReverseComplement(OverlapsList[i])
 RCFirstSequence = ReverseComplement(FirstSequence)

 #location of reverse complement of overlap in reverse complement of sequence
 #(index of last residue of overlap is OverlapIndex + (OverlapLength-1)
 #(inversion of index for base-0 indexing is x ---> N - x, therefore overlapRC = N - overlap - len(overlap) + 1 !!!
 RCOverlapIndex = len(FinalSequence) - (OverlapIndexList[i] + (len(OverlapsList[i]) - 1))
 
 RCPrimerStartIndex = RCOverlapIndex

 if RCFirstFragmentIndex > RCPrimerStartIndex:
  PrimerAnnealStartIndex = RCFirstFragmentIndex
 else:
  PrimerAnnealStartIndex = RCPrimerStartIndex
 
 #LEGACY CODE
 #if seq2r.find(RCFirstSequence) >= seq2r.find(l3ir):
 # bigger1 = seq2r.find(RCFirstSequence)
 #else:
 # bigger1 = seq2r.find(l3ir)
  
 for tPrimerLength in range(MinPrimerAnneal,MaxPrimerAnneal + 1):

  FragmentEndIndex = RCFirstFragmentIndex + tPrimerLength
  TerminalResidue = RCFinalSequence[FragmentEndIndex].upper()
  tAnnealSequence = RCFinalSequence[PrimerAnnealStartIndex:FragmentEndIndex + 1]
  tempTm = DNA_Tm(tAnnealSequence,DefaultSalt)

  if (tempTm >= minPrimerTm) and (TerminalResidue == 'C' or TerminalResidue == 'G'):      
   break


 print
 
 PrintTm_rev = str(tempTm)[:4]
 PrintSeq_rev = RCFinalSequence[RCPrimerStartIndex:FragmentEndIndex + 1]
 
 print '>Overlap ' + str(i + 1) + ': ' + TemplateName1 + '_r  (Tm ' + PrintTm_rev + ')'
 print PrintSeq_rev
  
 #"""###final tagged output that is read by Gibson Assembly excel macro"""
 print '[OVERLAP][OverlapSequence]' + OverlapsList[i] + '[\OverlapSequence] [dG]' + str(FinalEnergyList[i]) + '[\dG] [Tm]' + str(TmList[i])[:4] + '[\Tm]'
 print '[PRIMER1][PrimerName]' + TemplateName2 + '_f[\PrimerName] [Sequence]' + PrintSeq_for + '[\Sequence] [Tm]' + PrintTm_for + '[\Tm]'
 print '[PRIMER2][PrimerName]' + TemplateName1 + '_r[\PrimerName] [Sequence]' + PrintSeq_rev + '[\Sequence] [Tm]' + PrintTm_rev + '[\Tm]'
 print
