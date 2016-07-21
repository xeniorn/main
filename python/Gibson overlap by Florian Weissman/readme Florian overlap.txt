The input.txt file should be filled in the following way:

Line 1: First sequence (part that is present on a template for PCR)
Line 2: Part that has to be built by oligos (e.g. Flag-tag, or part that introduces mutations, = everything not present on a PCR template)
Line 3: Second sequence (part that is present on a template for PCR)
Line 4: Number(s) of the lines which should not be used to find a Gibson homology sequence (line 4 can be empty, but has to be present in the file!)

Possible entries in line 4:
empty   Homology region will be found somewhere around the center (standard)
1               Homology region will be found in lines 2&3
2               Homology region will be found in line 1 or 3
3               Homology region will be found in lines 1&2
12              Homology region will be found in line 3
23              Homology region will be found in line 1
else            Homology region will be found somewhere around the center

Example:
42 different cDNAs are supposed to be cloned into the same vector. Because one does not want to look 42 times for good Gibson overhangs, high efficiency overhangs should be found within the vector sequence. This way the vector has to be prepared by PCR only once. Oligos to make the inserts carry the overhangs and can be very easily designed by changing the gene-specific parts.

For the N-terminus:
Line1: vector
Line2: maybe tag
Line3: cDNA1
Line4: 23

For the C-terminus:
Line1: cDNA1
Line2: maybe tag
Line3: vector
Line4: 12