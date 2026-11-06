# Midterm Presentation Script

# Load data
evals <- read.csv("/Users/admin/Desktop/DOC/DSS 665/evals.csv", header = TRUE)

head(evals)

summary(evals)

# Install once if needed
# install.packages("skimr")

# Load package
library(skimr)

# Full dataset overview
skim(evals)

###Change number of significant figures displayed (default = 3)
options(pillar.sigfig = 2)

skim(evals)

# Focus on one variable
skim(evals, score)

# Select multiple variables
skim(evals, score, age, bty_avg)

# Summary without mini charts
skim_without_charts(evals, score, age, bty_avg)

# Save skim output summary as an object
skim_results <- skim(evals)
class(skim_results)
View(skim_results)

