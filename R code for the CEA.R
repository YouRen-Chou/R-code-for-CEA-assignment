#####Import data from the dataset.xlsx#
#To make sure that R code could be imported currently, "readx1" package needs to be installed
# And also the dataset excel should be download on the desk.
install.packages("readx1")
library(readxl)
dataset <- read_xlsx(path = "Dataset.xlsx", range = "A1:L201")

#rename the column name
names(dataset)[2] <- "Group"
names(dataset)[5] <- "Response"
names(dataset)[7] <- "Intervention_costs"
names(dataset)[10] <- "Inpatient_costs"
names(dataset)[11] <- "Outpatient_costs"
names(dataset)[12] <- "Medication_costs"

#-------------------------------------------------------------------------------
##### Task No.1 ##### Define the Research Question.
#-------------------------------------------------------------------------------
#PICO question form set:
#P: Patients experiencing moderate and severe depression.
#I: Treatment X (assumed), could be either therapy, medication or a combination.
#C: Standard care.
#O: The intervention is cost-effective
    #(Gain more health benefit-QALYs)?
    #(The intervention is cheaper or more expensive-Costs)?
    #(Incremental cost-effective ratio- ICER)?
#To determine whether the intervention is cost-effective compared to standard care

# (Study Report. Table 1) Patient Characteristic
library(dplyr)
library(gtsummary)

dataset %>%
  mutate(
    Group = ifelse(Group == 0, "Control Group", "Treatment Group"),  # Change the group name
    Age_group = ifelse(Age < 65, "< 65 years", ">= 65 years")  # Create Age group
  ) %>%
  select(Age, Age_group, Sex, Group) %>%  # Reorder the variables here
  tbl_summary(
    by = Group,  # Group by control or treatment groups
    statistic = list(
      all_continuous() ~ "{mean} ({sd})",        # Show mean and sd
      all_categorical() ~ "{n} / {N} ({p}%)"     # Show count and percentage for categorical variables
    ),
    digits = all_continuous() ~ 1,  
    type   = all_categorical() ~ "categorical",  
    label  = list(  # Rename the labels
      Age ~ "All Patient age",
      Sex ~ "Gender",
      Age_group ~ "Age Group"
    )
  )

#-------------------------------------------------------------------------------
##### Task No.2 ##### Calculate Costs.
#-------------------------------------------------------------------------------
# Create a new column to show the total costs for each patient.
dataset$Total_costs_per_patients <- (dataset$Intervention_costs + 
                                       dataset$Inpatient_costs + 
                                       dataset$Outpatient_costs+ 
                                       dataset$Medication_costs )
print(n = 200, dataset[,"Total_costs_per_patients", drop=FALSE])
#N.B. Noted that the productivity losses are not included in this stage, they would be conducted through sensitivity analysis.


#-------------------------------------------------------------------------------
##### Task No.3 ##### Compare Costs and Outcomes Between Groups.
#-------------------------------------------------------------------------------
############### a. Compute mean total costs and mean QALYs for both groups.
#-------------------------------------------------------------------------------
library(dplyr)
library(tidyr)

# The mean of total costs by groups
dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(mean_total_costs = mean(Total_costs_per_patients))

# The mean of QAYLs by groups
dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(mean_QALYs = mean(QALYs))

####
# (Study Report. Figure 1a) Total Costs Forest plot
library(dplyr)
library(forestplot)
library(grid)

summary_stats <- dataset %>%
  group_by(Group, Sex) %>%
  summarise(
    Mean = mean(Total_costs_per_patients),
    SD = sd(Total_costs_per_patients),
    n = n(),
    Lower = Mean - 1.96 * (SD / sqrt(n)),  # calculate the 95% CI
    Upper = Mean + 1.96 * (SD / sqrt(n))
  ) %>%
  ungroup()

summary_total <- dataset %>%
  group_by(Group) %>%
  summarise(
    SEX = "Total",
    Mean = mean(Total_costs_per_patients),
    SD = sd(Total_costs_per_patients),
    n = n(),
    Lower = Mean - 1.96 * (SD / sqrt(n)),
    Upper = Mean + 1.96 * (SD / sqrt(n))
  )

summary_stats <- bind_rows(summary_stats, summary_total)
summary_stats <- summary_stats %>%
  mutate(Group = ifelse(Group == 0, "Control", "Intervention"))
summary_stats$Sex[is.na(summary_stats$Sex)] <- "Total"  # Replace the NA into total
ndx = order(summary_stats$Group, decreasing=F)          # Reorder the column by group
summary_stats_sorted = summary_stats[ndx,]              # Create the sorted data frame

# Creat data frame
data <- data.frame(
  Group = c("Control", "Treatment"),
  Sex = c(summary_stats_sorted$Sex),
  n = c(summary_stats_sorted$n),
  Costs_mean = c(summary_stats_sorted$Mean),
  Costs_lower = c(summary_stats_sorted$Lower),
  Costs_upper = c(summary_stats_sorted$Upper)
)

costs_colors <- fpColors(box = "blue", lines = "black", summary = "black") # Colour setting
mean_costs <- sprintf("%.2f", summary_stats_sorted$Mean) # Mean costs column setting

row_names <- cbind(c("Groups", "Control", "", "", "Treatment", "", ""),
                   c("Sex", "female", "male" , "Total", "female", "male" , "Total"),
                   c("No" , summary_stats_sorted$n),
                   c("Mean Costs", round(summary_stats_sorted$Mean,2))
)

forestplot(
  labeltext = row_names, 
  mean = c(NA, summary_stats_sorted$Mean), 
  lower = c(NA, summary_stats_sorted$Lower), 
  upper = c(NA, summary_stats_sorted$Upper), 
  is.summary = c(TRUE, FALSE, FALSE), 
  xlog = FALSE,
  xlim = c(7000, 11000),  # Adjust x-axis range
  xlab = "Mean Costs (with 95% CI)", 
  col = costs_colors, 
  boxsize = 0.2, 
  ci.vertices = TRUE,
  title = "Total Costs",
  align = c("l", "l", "c", "r"),  # Adjust column alignment
  graph.pos = 4  # Moves the plot closer to the text
)

####
# (Study Report. Figure 1b) Total QALYs Forest plot
library(dplyr)
library(forestplot)
library(grid)

summary_stats <- dataset %>%
  group_by(Group, Sex) %>%
  summarise(
    Mean = mean(QALYs),
    SD = sd(QALYs),
    n = n(),
    Lower = Mean - 1.96 * (SD / sqrt(n)),  # calculate the 95% CI
    Upper = Mean + 1.96 * (SD / sqrt(n))
  ) %>%
  ungroup()

summary_total <- dataset %>%
  group_by(Group) %>%
  summarise(
    SEX = "Total",
    Mean = mean(QALYs),
    SD = sd(QALYs),
    n = n(),
    Lower = Mean - 1.96 * (SD / sqrt(n)),
    Upper = Mean + 1.96 * (SD / sqrt(n))
  )

summary_stats <- bind_rows(summary_stats, summary_total)
summary_stats <- summary_stats %>%
  mutate(Group = ifelse(Group == 0, "Control", "Intervention"))
summary_stats$Sex[is.na(summary_stats$Sex)] <- "Total"  # Replace the NA into total
ndx = order(summary_stats$Group, decreasing=F)          # Reorder the column by group
summary_stats_sorted = summary_stats[ndx,]              # Create the sorted data frame

# Create data frame
data <- data.frame(
  Group = c("Control", "Treatment"),
  Sex = c(summary_stats_sorted$Sex),
  n = c(summary_stats_sorted$n),
  QALYs_mean = c(summary_stats_sorted$Mean),
  QALYs_lower = c(summary_stats_sorted$Lower),
  QALYs_upper = c(summary_stats_sorted$Upper)
)

QALYs_colors <- fpColors(box = "red", lines = "black", summary = "black") # Colour setting
mean_QALYs <- sprintf("%.2f", summary_stats_sorted$Mean) # Mean costs column setting

row_names <- cbind(c("Groups", "Control", "", "", "Treatment", "", ""),
                   c("Sex", "female", "male" , "Total", "female", "male" , "Total"),
                   c("No" , summary_stats_sorted$n),
                   c("Mean QALYs", round(summary_stats_sorted$Mean,2))
)

forestplot(
  labeltext = row_names, 
  mean = c(NA, summary_stats_sorted$Mean), 
  lower = c(NA, summary_stats_sorted$Lower), 
  upper = c(NA, summary_stats_sorted$Upper), 
  is.summary = c(TRUE, FALSE, FALSE), 
  xlog = FALSE,
  xlim = c(0.7,1.0),  # Adjust x-axis range
  xlab = "Mean QALYs (with 95% CI)", 
  col = QALYs_colors, 
  boxsize = 0.2, 
  ci.vertices = TRUE,
  title = "Total QALYs",
  align = c("l", "l", "c", "r"),  # Adjust column alignment
  graph.pos = 4  # Moves the plot closer to the text
)


#-------------------------------------------------------------------------------
############### b. Calculate the incremental cost and incremental effectiveness(both response and QALYs).
#-------------------------------------------------------------------------------
library(dplyr)
library(tidyr)

# Incremental Costs
Incremental_cost_table <- dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(Mean_total_costs = mean(Total_costs_per_patients, na.rm = TRUE)) %>%  
  pivot_wider(names_from = Group, values_from = Mean_total_costs) %>% 
  mutate(Incremental_costs = `Treatment Group` - `Control Group`)

print(Incremental_cost_table)

# Incremental QALYs
Incremental_QALYs_table <- dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(Mean_QALYs = mean(QALYs, na.rm = TRUE)) %>%  
  pivot_wider(names_from = Group, values_from = Mean_QALYs) %>%  
  mutate(Incremental_QALYs = `Treatment Group` - `Control Group`)

print(Incremental_QALYs_table)

# Incremental treatment response
Incremental_treatment_response_table <- dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(Mean_Response = mean(Response, na.rm = TRUE)) %>%  
  pivot_wider(names_from = Group, values_from = Mean_Response) %>%  
  mutate(Incremental_Response = `Treatment Group` - `Control Group`)

print(Incremental_treatment_response_table)


#-------------------------------------------------------------------------------
##### Task No.4 ##### Calculate the ICER.
#-------------------------------------------------------------------------------
library(dplyr)
library(tidyr)

ICER_table <- dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(
    Mean_Costs = mean(Total_costs_per_patients),  
    Mean_QALYs = mean(QALYs)  
    ) %>%
  pivot_wider(names_from = Group, values_from = c(Mean_Costs, Mean_QALYs)) %>%
  mutate(
    Incremental_Costs = `Mean_Costs_Treatment Group` - `Mean_Costs_Control Group`,
    Incremental_QALYs = `Mean_QALYs_Treatment Group` - `Mean_QALYs_Control Group`,
    ICER = Incremental_Costs / Incremental_QALYs
     )
value <- ICER_table[["ICER"]]
print(value)


#-------------------------------------------------------------------------------
##### Task No.5 ##### Compare the results with Sweden's WTP Threshold.
#-------------------------------------------------------------------------------
library(ggplot2)

icer_data <- data.frame(
  Treatment = c("Treatment X"),
  Incremental_Costs = c(ICER_table[["Incremental_Costs"]]), 
  Incremental_QALYs = c(ICER_table[["Incremental_QALYs"]])    
  )
WTP_threshold <- 500000
#Sweden's WTP threshold = 500,000 SEK per QALY. It is flexible to chang WTP value from here.

# Plot the ICER plane
ggplot(icer_data, aes(x = Incremental_QALYs, y = Incremental_Costs)) +
  geom_point(size = 3, color = "blue") +  # draw the ICER point
  geom_text(aes(label = Treatment), vjust = -1, hjust = 1, size = 5) + # add label on ICER point
  geom_abline(slope = WTP_threshold, intercept = 0, linetype = "dashed", color = "red") +  # WTP threshold 
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray") +  # x axis <- costs
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray") +  # y axis <- QALYs
  labs(
    title = "Cost-effectiveness Plane (Sweden’s WTP threshold)",
    x = "Incremental QALYs",
    y = "Incremental Costs (SEK)"
  ) +
  xlim(-0.1, 0.1) +  # adjust the range of x axis
  ylim(-2500, 5000) +  # adjust the range of y axis
  theme_minimal()	


#-------------------------------------------------------------------------------
##### Task No.6 ##### Conduct Sensitivity Analysis
#-------------------------------------------------------------------------------
#Include productivity losses (both presenteeism and absenteeism costs) in the cost-effectiveness model
# Create a new column to show the total costs which including productivity losses for each patient.
dataset$Total_costs_SA <- ( dataset$Total_costs_per_patients + 
                            dataset$Presenteism + 
                            dataset$Absenteism
                            )
print(n = 200, dataset[,"Total_costs_SA", drop=FALSE])

# Calculate the ICER based on the new total costs
ICER_SA_table <- dataset %>%
  mutate(Group = ifelse(Group == "0", "Control Group", "Treatment Group")) %>%
  group_by(Group) %>%
  summarise(
    Mean_Costs = mean(Total_costs_SA),  
    Mean_QALYs = mean(QALYs)  
  ) %>%
  pivot_wider(names_from = Group, values_from = c(Mean_Costs, Mean_QALYs)) %>%
  mutate(
    Incremental_Costs = `Mean_Costs_Treatment Group` - `Mean_Costs_Control Group`,
    Incremental_QALYs = `Mean_QALYs_Treatment Group` - `Mean_QALYs_Control Group`,
    ICER = Incremental_Costs / Incremental_QALYs
  )
value <- ICER_SA_table[["ICER"]]
print(value)

# Plot the ICER plane
library(ggplot2)

icer_SA_data <- data.frame(
  Treatment = c("Intervention"),
  Incremental_Costs = c(ICER_SA_table[["Incremental_Costs"]]), 
  Incremental_QALYs = c(ICER_SA_table[["Incremental_QALYs"]])    
)

ggplot(icer_SA_data, aes(x = Incremental_QALYs, y = Incremental_Costs)) +
  geom_point(size = 3, color = "darkorange") +  # draw the ICER point
  geom_text(aes(label = Treatment), vjust = -1, hjust = 1, size = 5) + # add label on ICER point
  geom_abline(slope = WTP_threshold, intercept = 0, linetype = "dashed", color = "red") +  # WTP threshold 
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray") +  # x axis <- costs
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray") +  # y axis <- QALYs
  labs(
    title = "ICER Plane (Account for Productivity losses)",
    x = "Incremental QALYs",
    y = "Incremental Costs (SEK)"
  ) +
  xlim(-0.1, 0.1) +  # adjust the range of x axis
  ylim(-2500, 5000) +  # adjust the range of y axis
  theme_minimal()	

# Additional: comparing two ICER  value (with and without including productivity losses)
library(ggplot2)
icer_all_data <- data.frame(
  Treatment = c("Without Losses", "With Losses"),
  Incremental_Costs = c(ICER_table[["Incremental_Costs"]],
                        ICER_SA_table[["Incremental_Costs"]]), 
  Incremental_QALYs = c(ICER_table[["Incremental_QALYs"]],
                        ICER_SA_table[["Incremental_QALYs"]])
  )

ggplot(icer_all_data, aes(x = Incremental_QALYs, y = Incremental_Costs, color = Treatment)) +
  geom_point(size = 3) +  # ICER point
  geom_abline(slope = WTP_threshold, intercept = 0, linetype = "dashed", color = "red") +  # WTP Threshold
  geom_hline(yintercept = 0, linetype = "dashed", color = "gray") +  # x axis <- costs
  geom_vline(xintercept = 0, linetype = "dashed", color = "gray") +  # y axis <- QALYs
  scale_color_manual(values = c("blue", "darkorange")) +  # different colour
  labs(
    title = "ICER Plane: Comparing Treatments",
    x = "Incremental QALYs",
    y = "Incremental Costs (£)"
  ) +
  xlim(-0.1, 0.1) +  # adjust the range of x axis
  ylim(-2500, 5000) +  # adjust the range of y axis
  theme_minimal()
