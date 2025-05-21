# R包
library(readxl)
library(ggplot2)
library(dplyr)
library(scales)
library(tidyr)
library(ggpubr)
library(pheatmap)
library(reshape2)
library(VIM)

# Sample导入 ----
Sample <- read_excel("01_rawdata/Sample.xlsx")
View(Sample)
Sample <- data.frame(Sample)

# 统计分析 ----
## 年龄分布展示----


ggplot(Sample, aes(x = Age, fill = Gender)) +
  geom_histogram(binwidth = 1, position = "stack", color = "white") +
  scale_fill_manual(values = c("M" = "#669aba", "F" = "#be1420")) +
  labs(title = "Age Distribution by Gender",
       x = "Age", y = "Count", fill = "Gender") +
  theme_minimal(base_size = 14)
ggsave("Age Distribution by Gender.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

## 男女比例 ----
# 计算每组比例
plot_data <- Sample %>%
  group_by(AgeGroup, Gender) %>%
  summarise(n = n(), .groups = "drop") %>%
  group_by(AgeGroup) %>%
  mutate(prop = n / sum(n),
         label = percent(prop, accuracy = 1))  # 百分比标签

# 画图
setwd("./04_figure/")
AgeGroup_order <- c("Young", "Adult", "Elderly")
plot_data$AgeGroup <- factor(plot_data$AgeGroup, levels = AgeGroup_order)

ggplot(plot_data, aes(x = AgeGroup, y = prop, fill = Gender)) +
  geom_bar(stat = "identity", position = "stack") +
  geom_text(aes(label = label),
            position = position_stack(vjust = 0.5), size = 4, color = "white") +
  scale_y_continuous(labels = scales::percent_format(accuracy = 10)) +
  scale_fill_manual(values = c("M" = "#669aba", "F" = "#be1420")) + 
  ggtitle("Gender Proportion") +
  theme_minimal(base_size = 14) +
  theme(legend.position = "right",
        plot.title = element_text(hjust = 0.5))
ggsave("SexRatio_by_AgeGroup.pdf",
       width = 6, height = 8, dpi = 300, bg = "white")




## 各亚组血常规指标 ----
blood_vars <- c("WBC", "NEU_count", "LYM_count", "MONO_count", "EOS_count", "BASO_count", 
                "NEU_percent", "LYM_percent", "MONO_percent", "EOS_percent", "BASO_percent",
                "RBC", "HGB", "HCT", "PLT")
### 统计描述（均值、标准差、中位数、IQR）----
desc_stats <- Sample %>%
  group_by(SubGroup) %>%
  summarise(across(all_of(blood_vars), list(mean = ~mean(.x, na.rm=TRUE),
                                            sd = ~sd(.x, na.rm=TRUE),
                                            median = ~median(.x, na.rm=TRUE),
                                            IQR = ~IQR(.x, na.rm=TRUE))))
print(desc_stats)

long_data <- Sample %>%
  select(SubGroup, all_of(blood_vars)) %>%
  pivot_longer(cols = -SubGroup, names_to = "Indicator", values_to = "Value")

# 先定义因子顺序
subgroup_levels <- c("Young_1", "Young_2", "Young_3", "Young_4",
                     "Adult_1", "Adult_2", "Adult_3", "Adult_4",
                     "Elderly_1", "Elderly_2", "Elderly_3")

# 在 long_data 中设置 SubGroup 为因子并指定顺序
long_data$SubGroup <- factor(long_data$SubGroup, levels = subgroup_levels)


ggplot(long_data, aes(x = SubGroup, y = Value, fill = SubGroup)) +
  geom_boxplot(outlier.shape = NA) +
  geom_jitter(width = 0.2, alpha = 0.4, size = 0.5) +
  facet_wrap(~ Indicator, scales = "free_y", ncol = 4) +
  theme_minimal(base_size = 12) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
  labs(title = "Blood routine indicators distribution by SubGroup", x = "SubGroup", y = "Number")
ggsave("Blood routine indicators distribution by SubGroup.pdf",
       width = 12, height = 9, dpi = 300, bg = "white")

### 按AgeGroup统计最大值、最小值、均值
agegroup_stats <- Sample %>%
  group_by(AgeGroup) %>%
  summarise(across(all_of(blood_vars),
                   list(Min = ~min(.x, na.rm = TRUE),
                        Max = ~max(.x, na.rm = TRUE),
                        Mean = ~mean(.x, na.rm = TRUE)),
                   .names = "{.col}_{.fn}"))

### 按SubGroup统计最大值、最小值、均值
subgroup_stats <- Sample %>%
  group_by(SubGroup) %>%
  summarise(across(all_of(blood_vars),
                   list(Min = ~min(.x, na.rm = TRUE),
                        Max = ~max(.x, na.rm = TRUE),
                        Mean = ~mean(.x, na.rm = TRUE)),
                   .names = "{.col}_{.fn}"))
setwd("../03_result/")
write.xlsx(agegroup_stats, "AgeGroup_Blood_Stats.xlsx")
write.xlsx(subgroup_stats, "SubGroup_Blood_Stats.xlsx")

# 组间差异检验
# 正态分布假设成立时用ANOVA
# 不满足正态用Kruskal-Wallis非参数检验
# WBC
# 检查正态性


# 批量检验
library(broom)

results <- lapply(blood_vars, function(var){
  formula <- as.formula(paste(var, "~ SubGroup"))
  res <- kruskal.test(formula, data = Sample)
  tidy(res)
})

results_df <- do.call(rbind, results)
results_df$variable <- blood_vars
print(results_df)

library(openxlsx)
setwd("../03_result/")
write.xlsx(results_df, "SubGroup_BloodRoutine_KW_results.xlsx")

# 可视化
# 可视化前先按 p 值排序
results_df_plot <- results_df %>%
  arrange(p.value) %>%
  mutate(Significant = ifelse(p.value < 0.05, "Yes", "No"),
         variable = factor(variable, levels = variable))  # 保持原顺序

ggplot(results_df_plot, aes(x = variable, y = p.value, fill = Significant)) +
  geom_col() +
  geom_hline(yintercept = 0.05, linetype = "dashed", color = "red") +
  scale_fill_manual(values = c("Yes" = "#d62728", "No" = "#1f77b4")) +
  coord_flip() +  # 横向展示，变量多时更清晰
  labs(title = "Kruskal-Wallis Test P-values by Indicator",
       x = "Indicator",
       y = "p-value",
       fill = "Significant (p < 0.05)") +
  theme_minimal(base_size = 13)

setwd("../04_figure/")

Sample$AgeTier <- factor(Sample$AgeTier, levels = c(
  "0-2", "3-6", "7-11", "12-17", "18-30", "31-40", "41-50", "51-59", "60-69", "70-79", "≥80"
))

Sample$AgeGroup <- factor(Sample$AgeGroup, levels = c(
  "Young", "Adult", "Elderly"
))


### WBC----
ggplot(Sample, aes(x = AgeGroup, y = WBC, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "WBC (10^9/L)", title = "WBC Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("WBC Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")



### NEU_count----
ggplot(Sample, aes(x = AgeTier, y = NEU_count, fill = AgeTier)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(title = "NEU_count Distribution Across SubGroups",
       x = "Agetier", y = "NEU_count (10^9/L)") +
  ggtitle("NEU_count Distribution Across Agetier") +
  theme_minimal(base_size = 14) +
  theme(legend.position = "right",
        plot.title = element_text(hjust = 0.5))
ggsave("NEU_count Distribution Across AgeTier.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")



ggplot(Sample, aes(x = AgeGroup, y = NEU_count, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "NEU_count (10^9/L)", title = "NEU_count Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("NEU_count Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### RBC----
ggplot(Sample, aes(x = AgeGroup, y = RBC, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "RBC (10^12/L)", title = "RBC Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("RBC Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### LYM_count----
ggplot(Sample, aes(x = AgeGroup, y = LYM_count, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "LYM_count (10^9/L)", title = "LYM_count Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("LYM_count Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### MONO_count----
ggplot(Sample, aes(x = AgeGroup, y = MONO_count, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "MONO_count (10^9/L)", title = "MONO_count Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("MONO_count Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### EOS_count----
ggplot(Sample, aes(x = AgeGroup, y = EOS_count, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "EOS_count (10^9/L)", title = "EOS_count Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("EOS_count Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### BASO_count----
ggplot(Sample, aes(x = AgeGroup, y = BASO_count, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "BASO_count (10^9/L)", title = "BASO_count Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("BASO_count Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### NEU_percent----
ggplot(Sample, aes(x = AgeGroup, y = NEU_percent, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "NEU_percent (%)", title = "NEU_percent Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("NEU_percent Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### LYM_percent----
ggplot(Sample, aes(x = AgeGroup, y = LYM_percent, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "LYM_percent (%)", title = "LYM_percent Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("LYM_percent Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### MONO_percent----
ggplot(Sample, aes(x = AgeGroup, y = MONO_percent, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "MONO_percent (%)", title = "MONO_percent Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("MONO_percent Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### EOS_percent----
ggplot(Sample, aes(x = AgeGroup, y = EOS_percent, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "EOS_percent (%)", title = "EOS_percent Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("EOS_percent Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### BASO_percent----
ggplot(Sample, aes(x = AgeGroup, y = BASO_percent, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "BASO_percent (%)", title = "BASO_percent Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("BASO_percent Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### HGB ----
ggplot(Sample, aes(x = AgeGroup, y = HGB, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "HGB (g/L)", title = "HGB Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("HGB Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### HCT----
ggplot(Sample, aes(x = AgeGroup, y = HCT, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "HCT (%)", title = "HCT Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("HCT Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

### PLT----
ggplot(Sample, aes(x = AgeGroup, y = PLT, fill = AgeGroup)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, alpha = 0.4, color = "black", size = 1) +
  labs(x = "AgeGroup", y = "PLT (10^9/L)", title = "PLT Distribution Across AgeGroup") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5)) +
  stat_compare_means(comparisons = list(
    c("Young", "Adult"),
    c("Young", "Elderly"),
    c("Adult", "Elderly")
  ),
  method = "wilcox.test", label = "p.signif")
ggsave("PLT Distribution Across AgeGroup.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")

## 血常规均值热图----
# 先计算平均值
avg_data <- Sample %>%
  group_by(AgeGroup) %>%
  summarise(across(all_of(blood_vars), ~ mean(.x, na.rm = TRUE)))

# 转置为矩阵，指标为行，AgeGroup 为列
# 重新生成 heat_matrix（行：指标，列：AgeGroup）
heat_matrix <- as.matrix(t(avg_data[, -1]))
colnames(heat_matrix) <- avg_data$AgeGroup
rownames(heat_matrix) <- colnames(avg_data)[-1]

# 对每一行做 z-score 标准化
heat_matrix_scaled <- t(scale(t(heat_matrix)))  # 行标准化

# 绘制热图（颜色更鲜明）

pheatmap(heat_matrix_scaled,
         cluster_rows = TRUE,
         cluster_cols = FALSE,
         color = colorRampPalette(c("#2166ac", "#f7f7f7", "#b2182b"))(100),
         main = "Z-score Normalized Blood Indicators",
         fontsize_row = 10,
         fontsize_col = 10)

## 身高、体重、吸烟、饮酒缺失值情况----
colSums(is.na(Sample[, c("Height.cm.", "Weight.kg.", "Smoke", "Drunk")]))
sapply(Sample[, c("Height.cm.", "Weight.kg.", "Smoke", "Drunk")], function(x) {
  sum(is.na(x)) / length(x)
})

# 选择关注的变量
vars <- c("Height.cm.", "Weight.kg.", "Smoke", "Drunk")

# 将所有列统一转为字符型
Sample_char <- Sample %>%
  mutate(across(all_of(vars), as.character))

# 创建缺失矩阵
missing_summary <- Sample_char %>%
  select(SubGroup, all_of(vars)) %>%
  pivot_longer(cols = all_of(vars), names_to = "Variable", values_to = "Value") %>%
  mutate(Missing = ifelse(is.na(Value) | Value == "", "Missing", "Not Missing")) %>%
  group_by(SubGroup, Variable, Missing) %>%
  summarise(Count = n(), .groups = "drop")

ggplot(missing_summary, aes(x = SubGroup, y = Count, fill = Missing)) +
  geom_bar(stat = "identity", position = "fill") +
  facet_wrap(~ Variable, ncol = 2) +
  scale_y_continuous(labels = scales::percent) +
  scale_fill_manual(values = c("Not Missing" = "#4CAF50", "Missing" = "#F44336")) +
  labs(title = "Missing Data Proportion by SubGroup",
       x = "SubGroup", y = "Proportion", fill = "Status") +
  theme_minimal(base_size = 13) +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))
ggsave("NA in Height Weight Smoke Drunk.pdf",
       width = 8, height = 6, dpi = 300, bg = "white")


