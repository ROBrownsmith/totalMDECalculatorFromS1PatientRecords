# totalMDECalculatorFromS1PatientRecords
Excel VBA that takes information from a SystmOne(R) search and sums the Morphine Dose Equivalent (MDE) per patient from an array of their recent opioid prescriptions. MDE is calculated by obtaining the strength from the formulation description by regex and parsing the dose instructions to calculate how many dose units per day are prescribed.
