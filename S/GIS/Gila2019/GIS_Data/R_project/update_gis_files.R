# This script was used to update GIS files with new observations entered in Excel sheets. 
# To make the paths relative to the script location, it's intended to be run in RStudio.

# Import libraries to deal with Excel sheets, spatial data
library('xlsx')
library('tidyverse')
library('sf')

#------------------------------
# File I/O
#------------------------------
# Set working directory to script location
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# Read in Excel files
survey_xl <- read.xlsx(file = "../../DataEntry/SWFL/WIFL_Survey_Data_2019.xlsx", 
                       sheetName = "territories", 
                       colIndex = c(1:17),
                       as.data.frame = TRUE)

nests_xl <- read.xlsx(file = "../../DataEntry/SWFL/SWFL_Nests_2019.xlsx", 
                      sheetName = "Nests", 
                      colIndex = c(1:15),
                      as.data.frame = TRUE)

ybcu_xl <- read.xlsx(file = "../../DataEntry/YBCU/YBCU_surveys_2019.xlsx",
                     sheetName = "Observations",
                     colIndex = c(1:22),
                     as.data.frame = TRUE)

# Output geopackage locations
swflout <- "../SWFL_2019.gpkg"
ybcuout <- "../YBCU_2019.gpkg"

#------------------------------
# Prepare observation tables
#------------------------------
# Clean up empty rows
survey_xl <- filter(survey_xl, !is.na(Date))
nests_xl <- filter(nests_xl, !is.na(Nest))
ybcu_xl <- filter(ybcu_xl, !is.na(Date))

# Split YBCU into presence and absence
ybcu_presence <- filter(ybcu_xl, ybcu_xl$Presence.Absence == "P")
ybcu_absence <- filter(ybcu_xl, ybcu_xl$Presence.Absence == "A")

# Make sure coordinates are numeric
survey_xl$UTM.E <- as.numeric(survey_xl$UTM.E)
survey_xl$UTM.N <- as.numeric(survey_xl$UTM.N)

nests_xl$UTM_E <- as.numeric(nests_xl$UTM_E)
nests_xl$UTM_N <- as.numeric(nests_xl$UTM_N)

ybcu_presence$Projected.UTM.E <- as.numeric(ybcu_presence$Projected.UTM.E)
ybcu_presence$Projected.UTM.N <- as.numeric(ybcu_presence$Projected.UTM.N)

ybcu_absence$Wpt_UTM_E <- as.numeric(ybcu_absence$Wpt_UTM_E)
ybcu_absence$Wpt_UTM_N <- as.numeric(ybcu_absence$Wpt_UTM_N)

#------------------------------
# Create spatial datasets from tables
#------------------------------
# Plot coordinates spatially
survey_sf <- st_as_sf(x = survey_xl,
                      coords = c("UTM.E", "UTM.N"),
                      crs = "+proj=utm +zone=12 +ellps=GRS80 +datum=NAD83 +units=m +no_defs")

nests_sf <- st_as_sf(x = nests_xl,
                     coords = c("UTM_E", "UTM_N"),
                     crs = "+proj=utm +zone=12 +ellps=GRS80 +datum=NAD83 +units=m +no_defs")

ybcu_p_sf <- st_as_sf(x = ybcu_presence,
                      coords = c("Projected.UTM.E", "Projected.UTM.N"),
                      crs = "+proj=utm +zone=12 +ellps=GRS80 +datum=NAD83 +units=m +no_defs")

ybcu_a_sf <- st_as_sf(x = ybcu_absence,
                      coords = c("Wpt_UTM_E", "Wpt_UTM_N"),
                      crs = "+proj=utm +zone=12 +ellps=GRS80 +datum=NAD83 +units=m +no_defs")

#------------------------------
# Write to geopackages
#------------------------------
st_write(survey_sf, dsn = swflout, layer = "survey_data", update = TRUE, layer_options = "OVERWRITE=YES")
st_write(nests_sf, dsn = swflout, layer = "nests", update = TRUE, layer_options = "OVERWRITE=YES")

st_write(ybcu_p_sf,dsn = ybcuout, layer = "presence", update = TRUE, layer_options = "OVERWRITE=YES")
st_write(ybcu_a_sf,dsn = ybcuout, layer = "absence", update = TRUE, layer_options = "OVERWRITE=YES")
