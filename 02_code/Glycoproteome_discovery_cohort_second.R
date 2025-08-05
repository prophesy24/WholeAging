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
library(tidyverse)
library(naniar)
library(corrplot)
library(glmnet)
library(ggpubr)
library(openxlsx)

# Sample导入 ----
Sample <- read_excel("01_rawdata/Sample_glycoproteome_discovery_second.xlsx")
View(Sample)
Sample <- data.frame(Sample)

# 缺失值统计 ----
# 血常规指标列
blood_vars <- c("WBC", "NEU_count", "LYM_count", "MONO_count", "EOS_count", "BASO_count",
                "NEU_percent", "LYM_percent", "MONO_percent", "EOS_percent", "BASO_percent",
                "RBC", "HGB", "HCT", "PLT")

ggplot(Sample, aes(x = Age, fill = Gender)) +
  geom_histogram(binwidth = 1, position = "stack", color = "white") +
  scale_fill_manual(values = c("M" = "#669aba", "F" = "#be1420")) +
  scale_x_continuous(breaks = floor(min(Sample$Age)):ceiling(max(Sample$Age))) +
  labs(title = "Age Distribution by Gender",
       x = "Age", y = "Count", fill = "Gender") +
  theme_minimal(base_size = 14) +
  theme(
    axis.text = element_text(color = "black", size = 10),
    axis.title = element_text(color = "black"),
    plot.title = element_text(hjust = 0.5),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1)  # x轴刻度旋转
  )

# 保存图像：适当拉宽
ggsave("./04_figure/Glycoproteome_discovery_cohort/Age Distribution by Gender - Every Year Tick.pdf",
       width = 16, height = 6, dpi = 300, bg = "white")



### 年龄的最大值、最小值等----
# 整体
summary(Sample$Age)
mean(Sample$Age, na.rm = TRUE)
range(Sample$Age, na.rm = TRUE)

# 男女
Sample %>%
  group_by(Gender) %>%
  summarise(
    Min = min(Age, na.rm = TRUE),
    Q1 = quantile(Age, 0.25, na.rm = TRUE),
    Median = median(Age, na.rm = TRUE),
    Mean = mean(Age, na.rm = TRUE),
    Q3 = quantile(Age, 0.75, na.rm = TRUE),
    Max = max(Age, na.rm = TRUE),
    SD = sd(Age, na.rm = TRUE),
    Count = n()
  )

# 密度图
ggplot(Sample, aes(x = Age, fill = Gender)) +
  geom_density(alpha = 0.4) +
  scale_fill_manual(values = c("M" = "#669aba", "F" = "#be1420")) +
  labs(title = "Age Density by Gender", x = "Age", y = "Density") +
  theme_minimal(base_size = 14) +
  theme(plot.title = element_text(hjust = 0.5))


## 男女比例 ----
# 计算每组比例
plot_data <- Sample %>%
  group_by(AgeGroup, Gender) %>%
  summarise(n = n(), .groups = "drop") %>%
  group_by(AgeGroup) %>%
  mutate(prop = n / sum(n),
         label = percent(prop, accuracy = 1))  # 百分比标签

# 画图
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
ggsave("./04_figure/Glycoproteome_discovery_cohort/SexRatio_by_AgeGroup.pdf",
       width = 6, height = 8, dpi = 300, bg = "white")

# 指定比较组
comparisons <- list(c("M", "F"))

# 绘图
p <- ggplot(Sample, aes(x = Gender, y = Age, fill = Gender)) +
  geom_boxplot(outlier.shape = NA, alpha = 0.7) +
  geom_jitter(width = 0.2, size = 1, alpha = 0.4, color = "black") +
  scale_fill_manual(values = c("M" = "#4C78A8", "F" = "#F18FB1")) +
  labs(title = "Age Distribution by Gender",
       x = "Gender", y = "Age") +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5),
    axis.text = element_text(color = "black"),
    axis.title = element_text(color = "black")
  ) +
  stat_compare_means(
    comparisons = comparisons,
    method = "wilcox.test",
    label = "p.signif",      # 显示显著性符号，如 ns、*、**
    tip.length = 0.02        # 调整连线长度
  )

# 显示图
print(p)
ggsave("./04_figure/Glycoproteome_discovery_cohort/Age_Boxplot_by_Gender.pdf", p, width = 6, height = 6, dpi = 300, bg = "white")


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
ggsave("./04_figure/Glycoproteome_discovery_cohort/Blood routine indicators distribution by SubGroup.pdf",
       width = 12, height = 9, dpi = 300, bg = "white")

### 按AgeGroup统计最大值、最小值、均值----
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

write.xlsx(agegroup_stats, "./03_result/Glycoproteome_discovery/AgeGroup_Blood_Stats.xlsx")
write.xlsx(subgroup_stats, "./03_result/Glycoproteome_discovery/SubGroup_Blood_Stats.xlsx")

# 组间差异检验
# 正态分布假设成立时用ANOVA
# 不满足正态用Kruskal-Wallis非参数检验
# WBC
# 检查正态性


# 批量检验
results <- lapply(blood_vars, function(var){
  formula <- as.formula(paste(var, "~ SubGroup"))
  res <- kruskal.test(formula, data = Sample)
  tidy(res)
})

results_df <- do.call(rbind, results)
results_df$variable <- blood_vars
print(results_df)

write.xlsx(results_df, "./SubGroup_BloodRoutine_KW_results.xlsx")

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

Sample$SubGroup <- factor(Sample$SubGroup, levels = c("Young_1", "Young_2", "Young_3", "Young_4",
                                                      "Adult_1", "Adult_2", "Adult_3", "Adult_4",
                                                      "Elderly_1", "Elderly_2", "Elderly_3"))

# 遍历每个指标并绘图
for (var in blood_vars) {
  p <- ggplot(Sample, aes(x = SubGroup, y = .data[[var]], fill = AgeGroup)) +
    geom_boxplot(outlier.shape = NA, alpha = 0.7) +
    geom_jitter(width = 0.2, alpha = 0.4, size = 0.8, color = "black") +
    labs(title = paste(var, "Distribution Across SubGroups"),
         x = "SubGroup", y = var) +
    theme_minimal(base_size = 13) +
    theme(
      axis.text.x = element_text(angle = 45, hjust = 1),
      plot.title = element_text(hjust = 0.5),
    ) +
    scale_fill_brewer(palette = "Set2")
  
  # 显示图像
  print(p)
  
  # 可选：保存为 PDF
  ggsave(paste0("./04_figure/Glycoproteome_discovery_cohort/Blood_", var, "_by_SubGroup.pdf"),
         plot = p, width = 10, height = 6, dpi = 300)
}

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
ggplot(Sample, aes(x = SubGroup, y = PLT, fill = AgeGroup)) +
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

### 血常规组内差异 ----
result_list <- list()

for (var in blood_vars) {
  for (ag in unique(Sample$AgeGroup)) {
    
    sub_data <- Sample %>%
      filter(AgeGroup == ag, !is.na(.data[[var]])) %>%
      select(SubGroup, !!sym(var))
    
    if (length(unique(sub_data$SubGroup)) > 1) {
      combn_pairs <- combn(unique(sub_data$SubGroup), 2, simplify = FALSE)
      
      for (pair in combn_pairs) {
        group1 <- sub_data %>% filter(SubGroup == pair[1]) %>% pull(!!sym(var))
        group2 <- sub_data %>% filter(SubGroup == pair[2]) %>% pull(!!sym(var))
        
        if (length(group1) > 1 & length(group2) > 1) {
          p <- wilcox.test(group1, group2)$p.value
          result_list[[length(result_list) + 1]] <- data.frame(
            Indicator = var,
            AgeGroup = ag,
            Group1 = pair[1],
            Group2 = pair[2],
            p_value = p,
            Significance = ifelse(p < 0.001, "***",
                                  ifelse(p < 0.01, "**",
                                         ifelse(p < 0.05, "*", "ns")))
          )
        }
      }
    }
  }
}

# 合并为一个数据框
final_result <- bind_rows(result_list)

write.xlsx(final_result, "SubGroup_Comparison_by_AgeGroup.xlsx", row.Names = FALSE)


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
colSums(is.na(Sample[, c("Height", "Weight", "Smoking_status", "Drinking_status")]))
sapply(Sample[, c("Height", "Weight", "Smoking_status", "Drinking_status")], function(x) {
  sum(is.na(x)) / length(x)
})

# 选择关注的变量
vars <- c("Height", "Weight", "Smoking_status", "Drinking_status")

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


## 相关性分析----
# 探索是否有强相关的指标
blood_corr <- c("Age", "WBC", "NEU_count", "LYM_count", "MONO_count", "EOS_count", "BASO_count", 
                "NEU_percent", "LYM_percent", "MONO_percent", "EOS_percent", "BASO_percent",
                "RBC", "HGB", "HCT", "PLT")
blood_data <- Sample[, blood_corr]
blood_cor <- cor(blood_data, use = "pairwise.complete.obs", method = "spearman")
pdf("./04_figure/Glycoproteome_discovery_cohort/Blood_Correlation_Heatmap.pdf", width = 7, height = 6)
corrplot(blood_cor,
         method = "color",
         type = "upper",
         addCoef.col = "black",
         tl.col = "black",
         tl.cex = 0.8,
         number.cex = 0.7,
         col = colorRampPalette(c("blue", "white", "red"))(200),
         mar = c(1,1,1,1))
dev.off()

# 男性和女性的年龄趋势图 ----

# 你要展示的血常规指标
markers <- blood_vars


long_df <- Sample %>%
  select(Age, Gender, all_of(markers)) %>%
  pivot_longer(cols = all_of(markers), names_to = "Marker", values_to = "Value")

# 控制小图顺序
long_df$Marker <- factor(long_df$Marker, levels = markers)

# 自定义颜色
color_palette <- c("M" = "#67A3BD",  # 深蓝色
                   "F" = "#D1626B")  # 深粉色

# 绘图
p <- ggplot(long_df, aes(x = Age, y = Value, color = Gender)) +
  geom_point(alpha = 0.3, size = 1.2) +
  geom_smooth(method = "loess", se = FALSE, size = 1.3, span = 1) +
  facet_wrap(~ Marker, scales = "free", ncol = 4) +
  scale_color_manual(values = color_palette) +
  labs(x = "Age", y = "Measurement", title = "Age-related Trends of Blood Markers by Gender") +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold"),
    strip.text = element_text(size = 13, face = "bold"),
    axis.text = element_text(color = "black"),
    axis.title = element_text(color = "black"),
    legend.title = element_text(face = "bold"),
    legend.text = element_text(size = 12),
    aspect.ratio = 1  # 关键：强制每个面板为正方形
  )

# 显示图形
print(p)

# 保存（每个小图大致正方形，整图较大）
ggsave("Gender_Age_Trends_SquarePanels.pdf", plot = p,
       width = 14, height = 14, dpi = 300, bg = "white")


# 将筛选的样本和available样本坐标轴对齐，看血常规趋势图----
markers <- blood_vars
Sample1 <- read_excel("01_rawdata/Sample_glycoproteome_discovery_second.xlsx")
View(Sample1)
Sample1 <- data.frame(Sample1)

Sample2 <- read_excel("01_rawdata/Sample_available.xlsx")
View(Sample2)
Sample2 <- data.frame(Sample2)

Sample1 <- Sample1 %>%
  mutate(GroupID = as.character(GroupID))

Sample2 <- Sample2 %>%
  mutate(GroupID = as.character(GroupID))

library(dplyr)

# 合并数据以获取全局范围
combined_data <- bind_rows(
  Sample1 %>% mutate(Source = "Data1"),
  Sample2 %>% mutate(Source = "Data2")
)

# 血常规指标列表
markers <- blood_vars

# 计算每个指标的全局最小最大值
marker_limits <- combined_data %>%
  pivot_longer(cols = all_of(markers), names_to = "Marker", values_to = "Value") %>%
  group_by(Marker) %>%
  summarise(
    ymin = min(Value, na.rm = TRUE),
    ymax = max(Value, na.rm = TRUE)
  )

library(ggplot2)

plot_facet_with_limits <- function(data, source_name) {
  long_df <- data %>%
    select(Age, Gender, all_of(markers)) %>%
    pivot_longer(cols = all_of(markers), names_to = "Marker", values_to = "Value") %>%
    mutate(Marker = factor(Marker, levels = markers))
  
  # 合并 y 限制信息
  plot_data <- left_join(long_df, marker_limits, by = "Marker")
  
  ggplot(plot_data, aes(x = Age, y = Value, color = Gender)) +
    geom_point(alpha = 0.3, size = 1.2) +
    geom_smooth(method = "loess", se = FALSE, size = 1.3, span = 1) +
    facet_wrap(~ Marker, scales = "free", ncol = 4) +
    scale_color_manual(values = c("M" = "#67A3BD", "F" = "#D1626B")) +
    coord_cartesian(ylim = NULL) +  # 启用手动 y 范围
    labs(title = paste0("Age Trends in ", source_name), x = "Age", y = "Value") +
    theme_minimal(base_size = 14) +
    theme(
      plot.title = element_text(hjust = 0.5, face = "bold"),
      strip.text = element_text(size = 13, face = "bold"),
      axis.text = element_text(color = "black"),
      axis.title = element_text(color = "black"),
      legend.title = element_text(face = "bold"),
      aspect.ratio = 1
    ) +
    # 为每个 facet 单独设置 y 限（通过 ggforce::facet_wrap_paginate() 可细分）
    facet_wrap(~ Marker, scales = "free", ncol = 4) +
    geom_blank(aes(y = ymin)) +  # 强制 y 轴包含下限
    geom_blank(aes(y = ymax))   # 强制 y 轴包含上限
}

p1 <- plot_facet_with_limits(Sample1, "Dataset 1")
p2 <- plot_facet_with_limits(Sample2, "Dataset 2")

ggsave("./04_figure/Trend_by_Marker_Data1.pdf", p1, width = 14, height = 14, dpi = 300)
ggsave("./04_figure/Trend_by_Marker_Data2.pdf", p2, width = 14, height = 14, dpi = 300)
