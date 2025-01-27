# ========================================================================
# Script Title: LMDI Decomposition for CA Manufacturing Sector Emissions
# Author: Vedant Sinha
# Email: vedant.sinha22@gmail.com
# Date Created: 2025-01-26
# Last Modified: 2025-01-26
# Description: This R script reproduces the Logarithmic Mean Divisia Index (LMDI) calculations, results and visuals 
#              behind the academic poster titled "Decomposition of Energy-related Carbon Emissions: California's Manufacturing Sector.
# 
# License Information:
#   - **Poster (OSF)**: The academic poster is licensed under the Creative Commons 
#     Attribution 4.0 International License (CC BY ATTRIBUTION 4.0 INTERNATIONAL). You are free to 
#     share, remix, and adapt the poster, provided you give appropriate 
#     credit. Commercial use is allowed under this license.
#   - **Code (GitHub)**: The code behind this poster is licensed under the 
#     MIT License. You are free to use, modify, and distribute the code, 
#     provided that you include the original copyright notice and license 
#     in all copies or substantial portions of the code. 
#     This code is provided "as is", without warranty of any kind.
#
#  R Packages Used:
#
#  Schauberger P, Walker A (2025). _openxlsx: Read, Write and Edit xlsx Files_. R package version 4.2.8,
#  <https://CRAN.R-project.org/package=openxlsx>.
#
#  Wickham H, Bryan J (2023). _readxl: Read Excel Files_. R package version 1.4.3,
#  <https://CRAN.R-project.org/package=readxl>.
#
#  Wickham H, François R, Henry L, Müller K, Vaughan D (2023). _dplyr: A Grammar of Data Manipulation_. R package
#  version 1.1.4, <https://CRAN.R-project.org/package=dplyr>.
#
#  Hadley Wickham (2007). Reshaping Data with the reshape Package. Journal of Statistical Software, 21(12), 1-20.
#  URL http://www.jstatsoft.org/v21/i12/.
#
#  Wickham H, Vaughan D, Girlich M (2024). _tidyr: Tidy Messy Data_. R package version 1.3.1,
#  <https://CRAN.R-project.org/package=tidyr>.
#
#
#
#
#



library(readxl)
library(dplyr)
library(reshape2)
library(tidyr)
library(openxlsx)




current_directory <- getwd()
carb_data_file_name <- "raw_data_carb.xlsx"
carb_data_file_path <- file.path(current_directory, carb_data_file_name)
df <- read_excel(carb_data_file_path, sheet = "final_filtered")
df_naics_mapping <- read_excel(carb_data_file_path, sheet = "naics_mapping")
df_mapped <- merge(df, df_naics_mapping, by = c("Sub Sector Level 1", "Sub Sector Level 2", "Sub Sector Level 3"), all.x = TRUE)
df_fuelefs_mapping <- read_excel(carb_data_file_path, sheet = "fuel_mapping")
df_mapped2 <- merge(df_mapped, df_fuelefs_mapping, by = c("Activity Subset"), all.x = TRUE)
df_mapped3 <- as.data.frame(lapply(df_mapped2, function(x) ifelse(x == "No Data", 0, x)))
colnames(df_mapped3) <- colnames(df_mapped2)
df_mapped3[, 9:30] <- lapply(df_mapped3[, 9:30], as.numeric)

str(df_mapped3)
df_mapped3 <- pivot_longer(df_mapped3, cols = 9:30, names_to = "Year", values_to = "value")

df_mapped3$energy_mmbtu <- df_mapped3$value/df_mapped3$co2_ef

df_mapped3 <- df_mapped3 %>%
  group_by(NAICS, Year,`Activity Subset`) %>%
  summarise(total_energy_mmbtu = sum(energy_mmbtu, na.rm = TRUE), .groups = "drop")

df_mapped3 <- df_mapped3[df_mapped3$NAICS != "Oil and gas extraction", ]

df_mapped3 <- df_mapped3 %>%
  rename(Fuel = `Activity Subset`)




us_eia_mecs_electricity_data_file_name <- "ElecConsumptionOverall.xlsx"
us_eia_mecs_electricity_data_file_path <- file.path(current_directory, us_eia_mecs_electricity_data_file_name)
df_electricity <- read_excel(us_eia_mecs_electricity_data_file_path, sheet = "elec_data",col_names=T)
df_electricity <- pivot_longer(df_electricity, cols = 2:16, names_to = "NAICS", values_to = "electricity_million_kwh")
df_electricity$electricity_million_kwh <- (df_electricity$electricity_million_kwh * 10^6)

df_electricity <- df_electricity %>%
  rename(total_energy_mmbtu = electricity_million_kwh)

df_electricity$Fuel <- 'Electricity'

df_mapped3 <- rbind(df_mapped3, df_electricity)

heavy_sectors <- c("Chemical manufacturing", "Petroleum and coal products manufacturing") 

df_mapped3$NAICS <- ifelse(df_mapped3$NAICS %in% heavy_sectors, "Heavy", "Light")

df_mapped3 <- df_mapped3 %>%
  group_by(NAICS, Year, Fuel) %>%
  summarise(total_energy_mmbtu = sum(total_energy_mmbtu, na.rm = TRUE), .groups = "drop")

unique(df_mapped3$Fuel)


gdp_data_file_name <- "CARealGDPChained2012.xlsx"
gdp_data_file_path <- file.path(current_directory, gdp_data_file_name)
df_gdp <- read_excel(gdp_data_file_path, sheet = "gdp",col_names=T)
df_gdp <- pivot_longer(df_gdp, cols = 2:26, names_to = "Year", values_to = "gdp")
df_gdp <- df_gdp %>%
  rename(NAICS = Sectors)
df_gdp <- df_gdp[df_gdp$NAICS != "Oil and gas extraction", ]
df_gdp$NAICS <- ifelse(df_gdp$NAICS %in% heavy_sectors, "Heavy", "Light")
df_gdp <- df_gdp %>%
  group_by(NAICS, Year) %>%
  summarise(gdp = sum(gdp, na.rm = TRUE), .groups = "drop")
df_gdp <- df_gdp[df_gdp$Year == 2000 | df_gdp$Year == 2021, ]
df_mapped3 <- df_mapped3[df_mapped3$Year == 2000 | df_mapped3$Year == 2021, ]
df_gdp <- df_gdp[df_gdp$Year == 2000 | df_gdp$Year == 2021, ]

ActivityDf <- df_gdp


df_fuelefs_mapping2 <- read_excel(carb_data_file_path, sheet = "ef_mapping_2")
CO2_DF <- df_mapped3
CO2_DF <- merge(CO2_DF, df_fuelefs_mapping2, by = c("Fuel",'Year'), all.x = TRUE)
CO2_DF$CO2 <- CO2_DF$co2_ef * CO2_DF$total_energy_mmbtu

CO2_DF$Fuel[CO2_DF$Fuel != "Electricity"] <- "Fossil Fuels"

CO2_DF <- CO2_DF %>%
  group_by(Year,Fuel,NAICS) %>%
  summarise(CO2 = sum(CO2, na.rm = TRUE), .groups = "drop")

df_mapped3$Fuel[df_mapped3$Fuel != "Electricity"] <- "Fossil Fuels"

df_mapped3 <- df_mapped3 %>%
  group_by(Year,Fuel,NAICS) %>%
  summarise(total_energy_mmbtu = sum(total_energy_mmbtu, na.rm = TRUE), .groups = "drop")


CO2_DF_tot <- CO2_DF %>%
  group_by(Year) %>%
  summarise(CO2 = sum(CO2, na.rm = TRUE), .groups = "drop")




sectors <- unique(CO2_DF$NAICS)
fuels <- unique(CO2_DF$Fuel)
years <- unique(df_gdp$Year)

for (i in 1){
  base_yr = years[[i]]
  next_yr = years[[i+1]]
  yearselected <- c(base_yr,next_yr)
  CO2_DF <- CO2_DF[CO2_DF$Year %in% yearselected,]
  Loads_loop <- ActivityDf
  Loads_loop <- Loads_loop[Loads_loop$Year %in% yearselected,]
  ActivityT <- subset(Loads_loop, Year == next_yr)
  AT <- sum(ActivityT$gdp)
  Activity0 <- subset(Loads_loop, Year == base_yr)
  A0 <- sum(Activity0$gdp)
  constant = log(AT/A0)
  sector_sum <- data.frame(results = numeric(0))
  for (s in sectors){
    CO2Sectors <- subset(CO2_DF,NAICS == s)
    fuel_sum <- data.frame(results = numeric(0))
    for (j in fuels) {
      CO2Fuels <- subset(CO2Sectors,Fuel == j)
      CO2T <- subset(CO2Fuels, Year == next_yr)
      Co2T <- sum(CO2T$CO2)
      CO20 <- subset(CO2Fuels, Year == base_yr)
      Co20 <- sum(CO20$CO2)
      Wp <- (Co2T - Co20)/(log(Co2T)-log(Co20))
      deltaActEntry <- Wp*constant
      fuel_sum <- rbind(fuel_sum, data.frame(results = deltaActEntry))
    }
    fuels_total <- sum(fuel_sum$results)
    sector_sum <- rbind(sector_sum, data.frame(results = fuels_total))
  }
  ActivityEffect <- sum(sector_sum$results,na.rm=T)
}

ActivityEffect

sector_sum <- data.frame(results = numeric(0))
fuel_sum <- data.frame(results = numeric(0))


ActivityDf2 <- ActivityDf %>%
  group_by(Year) %>%
  mutate(Share = gdp / sum(gdp))


for (i in 1){
  base_yr = years[[i]]
  next_yr = years[[i+1]]
  yearselected <- c(base_yr,next_yr)
  sector_sum <- data.frame(results = numeric(0))
  CO2_DF <- CO2_DF[CO2_DF$Year %in% yearselected,]
  Loads_loop <- ActivityDf2
  Loads_loop <- Loads_loop[Loads_loop$Year %in% yearselected,]
  for (s in sectors){
    CO2Sectors <- subset(CO2_DF, NAICS == s)
    print(s)
    Loads_loop_sector <- subset(Loads_loop, NAICS == s)
    ShareT <- subset(Loads_loop_sector, Year == next_yr)
    Share0 <- subset(Loads_loop_sector, Year == base_yr)
    IShareT <- sum(ShareT$Share)
    IShare0 <- sum(Share0$Share)
    constant <- log(IShareT/IShare0)
    fuel_sum <- data.frame(results = numeric(0))
    for (j in fuels) {
      CO2Fuels <- subset(CO2Sectors,Fuel == j)
      CO2T <- subset(CO2Fuels, Year == next_yr)
      Co2T <- sum(CO2T$CO2)
      CO20 <- subset(CO2Fuels, Year == base_yr)
      Co20 <- sum(CO20$CO2)
      Wp <- (Co2T - Co20)/(log(Co2T)-log(Co20))
      deltaStrucEntry <- Wp*constant
      fuel_sum <- rbind(fuel_sum, data.frame(results = deltaStrucEntry))
    }
    fuels_total <- sum(fuel_sum$results)
    sector_sum <- rbind(sector_sum, data.frame(results = fuels_total))
  }
  StructureEffect <- sum(sector_sum$results,na.rm=T)
} 

StructureEffect

sector_sum <- data.frame(results = numeric(0))
fuel_sum <- data.frame(results = numeric(0))

EIDF <- df_mapped3 %>%
  group_by(Year,NAICS) %>%
  summarise(total_energy_mmbtu = sum(total_energy_mmbtu, na.rm = TRUE), .groups = "drop")

EIDF <- merge(EIDF, ActivityDf, by = c("Year","NAICS"), all.x = TRUE)
EIDF$EnergyIntensity <- EIDF$total_energy_mmbtu/EIDF$gdp

for (i in 1) {
  base_yr = years[[i]]
  next_yr = years[[i+1]]
  yearselected <- c(base_yr,next_yr)
  sector_sum <- data.frame(results = numeric(0))
  CO2_DF <- CO2_DF[CO2_DF$Year %in% yearselected,]
  EILoop <- EIDF
  EILoop <- EILoop[EILoop$Year %in% yearselected,]
  for (s in sectors){
    CO2Sectors <- subset(CO2_DF, NAICS == s)
    EILoop_Sector <- subset(EILoop, NAICS == s)
    EIT <- subset(EILoop_Sector, Year == next_yr)
    EIO <- subset(EILoop_Sector, Year == base_yr)
    EiT <- sum(EIT$EnergyIntensity)
    Ei0 <- sum(EIO$EnergyIntensity)
    constant <- log(EiT/Ei0)
    fuel_sum <- data.frame(results = numeric(0))
    for (j in fuels) {
      CO2Fuels <- subset(CO2Sectors,Fuel == j)
      CO2T <- subset(CO2Fuels, Year == next_yr)
      Co2T <- sum(CO2T$CO2)
      CO20 <- subset(CO2Fuels, Year == base_yr)
      Co20 <- sum(CO20$CO2)
      Wp <- (Co2T - Co20)/(log(Co2T)-log(Co20))
      deltaIntensEffectEntry <- Wp*constant
      fuel_sum <- rbind(fuel_sum, data.frame(results = deltaIntensEffectEntry))
    }
    fuels_total <- sum(fuel_sum$results)
    sector_sum <- rbind(sector_sum, data.frame(results = fuels_total))
  }
  IntensityEffect <- sum(sector_sum$results,na.rm=T)
} 

IntensityEffect

EMixDF <- df_mapped3 %>%
  group_by(Year,Fuel,NAICS) %>%
  summarise(total_energy_mmbtu = sum(total_energy_mmbtu, na.rm = TRUE), .groups = "drop")

EMixDF <- EMixDF %>%
  group_by(Year, NAICS) %>%
  mutate(EnergyMix = total_energy_mmbtu / sum(total_energy_mmbtu))

# EMixDF_an_test <- EMixDF %>%
#   group_by(Fuel,Year) %>%
#   summarise(total_energy_mmbtu = sum(total_energy_mmbtu, na.rm = TRUE), .groups = "drop")
# 
# EMixDF_test <- EMixDF_an_test %>%
#   group_by(Year) %>%
#   mutate(EnergyMix = total_energy_mmbtu / sum(total_energy_mmbtu))

sector_sum <- data.frame(results = numeric(0))
fuel_sum <- data.frame(results = numeric(0))

for (i in 1){
  base_yr = years[[i]]
  next_yr = years[[i+1]]
  yearselected <- c(base_yr,next_yr)
  sector_sum <- data.frame(results = numeric(0))
  CO2_DF <- CO2_DF[CO2_DF$Year %in% yearselected,]
  MixLoop <- EMixDF
  MixLoop <- MixLoop[MixLoop$Year %in% yearselected,]
  for (s in sectors){
    CO2Sectors <- subset(CO2_DF, NAICS == s)
    MixLoop_Sector <- subset(MixLoop, NAICS == s)
    fuel_sum <- data.frame(results = numeric(0))
    for (j in fuels) {
      MixLoop_Fuel <- subset(MixLoop_Sector, Fuel == j)
      EMT <- subset(MixLoop_Fuel, Year == next_yr)
      EMO <- subset(MixLoop_Fuel, Year == base_yr)
      EmT <- sum(EMT$EnergyMix)
      Em0 <- sum(EMO$EnergyMix)
      constant <- log(EmT/Em0)
      CO2Fuels <- subset(CO2Sectors,Fuel == j)
      CO2T <- subset(CO2Fuels, Year == next_yr)
      Co2T <- sum(CO2T$CO2)
      CO20 <- subset(CO2Fuels, Year == base_yr)
      Co20 <- sum(CO20$CO2)
      Wp <- (Co2T - Co20)/(log(Co2T)-log(Co20))
      deltaEMixEntry <- Wp*constant
      fuel_sum <- rbind(fuel_sum, data.frame(results = deltaEMixEntry))
    }
    fuels_total <- sum(fuel_sum$results)
    sector_sum <- rbind(sector_sum, data.frame(results = fuels_total))
  }
  EMixEffect <- sum(sector_sum$results,na.rm=T)
}

EMixEffect

EFDF <- merge(CO2_DF, df_mapped3, by = c("Year","NAICS","Fuel"), all.x = TRUE)
EFDF$EFValue <- EFDF$CO2/EFDF$total_energy_mmbtu

sector_sum <- data.frame(results = numeric(0))
fuel_sum <- data.frame(results = numeric(0))

for (i in 1){
  base_yr = years[[i]]
  next_yr = years[[i+1]]
  yearselected <- c(base_yr,next_yr)
  sector_sum <- data.frame(results = numeric(0))
  CO2_DF <- CO2_DF[CO2_DF$Year %in% yearselected,]
  EFLoop <- EFDF
  EFLoop <- EFLoop[EFLoop$Year %in% yearselected,]
  for (s in sectors){
    EFLoop_Sector <- subset(EFLoop, NAICS == s)
    CO2Sectors <- subset(CO2_DF, NAICS == s)
    fuel_sum <- data.frame(results = numeric(0))
    for (j in fuels) {
      EFLoop_Fuel <- subset(EFLoop_Sector, Fuel == j)
      EFT <- subset(EFLoop_Fuel, Year == next_yr)
      EFO <- subset(EFLoop_Fuel, Year == base_yr)
      EfT <- sum(EFT$EFValue)
      Ef0 <- sum(EFO$EFValue)
      constant <- log(EfT/Ef0)
      CO2Fuels <- subset(CO2Sectors,Fuel == j)
      CO2T <- subset(CO2Fuels, Year == next_yr)
      Co2T <- sum(CO2T$CO2)
      CO20 <- subset(CO2Fuels, Year == base_yr)
      Co20 <- sum(CO20$CO2)
      Wp <- (Co2T - Co20)/(log(Co2T)-log(Co20))
      deltaEFactorEntry <- Wp*constant
      fuel_sum <- rbind(fuel_sum, data.frame(results = deltaEFactorEntry))
    }
    fuels_total <- sum(fuel_sum$results)
    sector_sum <- rbind(sector_sum, data.frame(results = fuels_total))
  }
  EmissionsFactorEffect <- sum(sector_sum$results,na.rm=T)
}

EmissionsFactorEffect

total <- EmissionsFactorEffect + ActivityEffect + EMixEffect + IntensityEffect + StructureEffect
total

EmissionsFactorEffect
ActivityEffect
EMixEffect
IntensityEffect
StructureEffect

df_export <- data.frame(
  EmissionsFactorEffect = EmissionsFactorEffect,
  ActivityEffect = ActivityEffect,
  EMixEffect = EMixEffect,
  IntensityEffect = IntensityEffect,
  StructureEffect = StructureEffect
)

StructureEffect

df_gdp_summary <- df_gdp %>%
  group_by(Year) %>%
  summarise(gdp = sum(gdp, na.rm = TRUE), .groups = "drop")


lmdi_results_file_name <- "lmdi_results.xlsx"
lmdi_results_data_file_path <- file.path(current_directory, lmdi_results_file_name)

write.xlsx(df_export, file = lmdi_results_data_file_path, sheetName = "Sheet1")


CO2_DF_2000 <- CO2_DF[CO2_DF$Year == 2000, ]

CO2_DF_2021 <- CO2_DF[CO2_DF$Year == 2021, ]

sum(CO2_DF_2000$CO2)
sum(CO2_DF_2021$CO2)

