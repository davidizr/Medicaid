Quick Guide on this project and its structure:

# Dependencies not included in the github:

medicaid-provider-spending.csv - Download off medicaid data website
python packages: requirements.txt

To install packages do pip install -r requirements.txt


# Workflow 

Where the current CPT Codes + descriptions live: Codes Manual.xlsx
to add new codes, open that excel and add it to the list. There is conditional formatting to ensure no duplicates.

After new CPT codes are added, run Dataset_builder.ipynb. This creates a pivoted month by month table used for analysis. This process is memory intensive as the dataset is large. Consider closing all additional applications to free up memory. This job can take up to 10 minutes.

This job also contains the cutoff point of Dec 2023. 


Build slides of Poisson, NB, and Poisson Rates: Build_slides.ipynb
the backing code is in build_slides.py

In there, we have the Enrolees to calculate normalizaiton by 100k (interpolation of Apr 2022), the breakpoint (march 2021), dropping of March-May 2020, and all the logic to build the slides




