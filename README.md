# totalMDECalculatorFromS1PatientRecords
Excel VBA that takes information from a SystmOne(R) search and sums the Morphine Dose Equivalent (MDE) per patient from an array of their recent opioid prescriptions. MDE is calculated by obtaining the strength from the formulation description by regex and parsing the dose instructions to calculate how many dose units per day are prescribed. The array is sorted in descending MDE to allow quick identification of people receiveing >120MDE.

How to use:
Run the SystmOne search with an adhoc output as described below and save the output to ShowPatients.csv
Upon running the VBA it will ask you to specify the folder your search is in, then  all the data from the search will be imported to the workbook. NB a patient with a blank NHS number will currently break the macro.
The array formula is processor intensive and will take some time for practices with a large patient list.
After  you see  "Task Complete!" the MDE will be ranked largest to smallest, expand the selection.
Use caution in cases where MDE is zero or MANUAL or dose was not parsed.


Search format
medication\actiongroup - opioid analgesics
\report on earliest  matching issue of each applicable 
\report on all issues
\report on start date after 90 days ago

Ad hoc  output
Demographics\nhs number
Demographics\full  name
Registration Details\ Usual GP full name
Drug - include latest 9 values, action group opioid analgesics
Drug\Consultation date
Drug\medication  name
Drug\dose


Notes:
This is a work in progress.  Expected to be used by Pharmacy professionals in UK General practices. 
Users MUST check any calculations manually using own clinical knowledge. 
Conversion factors have been taken from the faculty of pain web site and these may change.  https://www.fpm.ac.uk/opioids-aware-structured-approach-opioid-prescribing/dose-equivalents-and-changing-opioids 
I would welcome collaboration in refining any of the sub functions, especially dose parsing. 

Current Limitations
Doses are calculated on what the maximum frequency they could be per day. Any dose that contains an "every 4 hours" component  is therefore assumed to be 6 times per day. This means  co-paracetamol item dose frequencies  are overstated when parsed. This is relevant for patients on multiple opiods close to MDE 120mg and warrants closer scrutiny.

Not every product is caught by the logic of the formulae yet. This may mean an abnormally large MDE is calculated or not at all.  Both cases warrant further investigation of the patient's notes.

If a dose instruction has a sequence of full stops to allow insertion of frequency or number of dose units, this can sometimes cause a problem. e.g. take...mL every...hour(s). Regex can seek and eliminate sequential full stops but if the user enters ther number after the first full stop it is difficult to programatically distinguish between take.5..mL and an intentional (but erroneously lacking preceding zero) take .5mL



