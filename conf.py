### For DVH Eclipse exports
# This variable is a list of all of the values that need to be extracted
DVH_VARIABLES = ['Min Dose [cGy]',
                 'Max Dose [cGy]',
                 'Mean Dose [cGy]',
                 'Modal Dose [cGy]',
                 'Median Dose [cGy]',
                 'STD [cGy]']
# These are other variables that are not as straightforward to get. These do not need to be changed, but are here for the sake of consistency
DVH_OTHER_VARIABLES = ['Fraction','Structure','DVH']

####################################################
####################################################

### FOR RT Excel sheets
# This variable is a list of all of the values that need to be extracted from the excel sheet
VARIABLES = ['date',
             'EBRT; dose per fraction',
             'EBRT; fractions without central shield',
             'EBRT; fractions with central shield',
             'EBRT; total dose',
             'BT; MR / CT',
             'BT; Applicator(s): type',
             'BT; Applicator(s): dimensions',
             'BT; prescribed dose PD',
             'BT; volume of PD [cm3]',
             'BT; volume of PDx2 [cm3]',
             'BT; pres. point level',
             'BT; pres. point [mm left',
             'BT; dose to + A left',
             'BT; dose to - A right',
             'BT; dose to A mean',
             'GTV  [cm3]',
             'GTV; D 100 = MTD',
             'GTV; D 90',
             'GTV; V 100',
             'HR CTV  [cm3]',
             'HR CTV; D 100 = MTD',
             'HR CTV; D 90',
             'HR CTV; V 100',
             'BLADDER  [cm3]',
             'BLADDER; ICRU - dose',
             'BLADDER; ICRUcr1,5cm - dose',
             'BLADDER; ICRUcr2,0cm - dose',
             'BLADDER; 0,1cm3 - dose',
             'BLADDER; 1cm3 - dose',
             'BLADDER; 2cm3 - dose',
             'RECTUM  [cm3]',
             'RECTUM; ICRU - dose',
             'RECTUM; ICRUprobe - dose',
             'RECTUM; 0,1cm3 - dose',
             'RECTUM; 1cm3 - dose',
             'RECTUM; 2cm3 - dose',
             'SIGMOID  [cm3]',
             'SIGMOID; 0,1cm3 - dose',
             'SIGMOID; 1cm3 - dose',
             'SIGMOID; 2cm3 - dose',
             'VAGINAL WALL; dose per fraction',
             'VAGINAL WALL; 1cm3 - dose',
             'VAGINAL WALL; 2cm3 - dose',
             'VAGINAL WALL; 5cm3 - dose',
             'VAGINAL WALL; 10cm3 - dose']

# This variable holds a dictionary where the keys are the section headers/structure names and the values are lists of different names that might be found in user-input corresponding to the keys.
TITLES = {		'PATIENT': ['Patient'],
				'EXTERNAL BEAM THERAPY': ['EBRT', 'External beam therapy'],
				'BRACHYTHERAPY': ['BT','Brachytherapy'],
				'GTV': ['GTV'],
				'PTV': ['PTV'],
				'HR CTV': ['HR CTV', 'HRCTV'],
				'IR CTV': ['IR CTV', 'IRCTV'],
				'INTESTINES': ['Intestines'],
				'BLADDER': ['Bladder'],
				'RECTUM': ['Rectum'],
				'SIGMOID': ['Sigmoid'],
				'VAGINAL WALL': ['Vaginal wall']}

# Controls some print statements for debug purposes
DEBUG = True

# Whatever you want to use as the delimiter between structure and value you want to extract
SPLICE = '; '