1. To install:
install.packages("devtools")  # Install devtools if not already installed in your R
devtools::install_github("EdKieran27/MRIOTs_Helper_Functions")

2. After installation, load the package:
library(MRIOTHelperFunctions)

3. To use the function for collapsing 72 economy MRIOT to 62 economy MRIOT, input three parameters: \n
   a. the directory where the 72 economy MRIOT is located
   b. the directory where you want to save the collapsed 62 economy MRIOT
   c. the year that you want to process

IMPORTANT: Make sure the MRIOT filename is structured like this: ADB-MRIO-2017.xlsx
Please rename the MRIOT file you want to compress like this for now, you can always change afterwards. 

sample:
aggregate_MRIO("C:/path/to/input", "C:/path/to/output", 2017)

