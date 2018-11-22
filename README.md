# final696Submit

Note: I had to make this repo last minute to submit the project because I messed up some things in git and almost erased all my work.
Here are the two previous git repos from which I built the final project. 
1) https://github.com/saldilf/696AuSpecProj.git  ==> I built >95% of the program functionality in this repo
2) https://github.com/saldilf/AuSpectra696.git   ==> I built the unit tests in this repo. 

This program does a few things that will save me a lot of time in the future. 

1) Reads in absorbance data for gold nanoparticles from an excel file and plots the data (abs vs wavelength). It can read from multiple columns. 
2) Extracts key parameters from each spectrum such as maximum absorbance, wavelength at max absorbance, and size of particle. 
3) Outputs the plots from (1) and a table summarizing the data in (2)

-----------
Things I would like to add to this program in the future:
  1) User-controlled visualization options. e.g. "do you want subplots or just one plot, or plot curves X, Y, Z on one plot and A, B, C on another," etc.
  2) Calculate concentration based on user input of extinction coefficinent and vial path length
  3) Let the user control what range of wavelengths to plot
  4) Let user decide how the legend looks based on how many columns are in the excel file
  5) Let user specify number of sheets in excel file and which ones to analyze
