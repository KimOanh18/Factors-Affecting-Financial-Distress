install.packages("readxl")
library(readxl)

data <- read_excel("K214142079.xlsx")
#Tính toán các biến nghiên cứu
data$Size <-log(data$`TỔNG TÀI SẢN`)
data$Current_ratio <- data$`TÀI SẢN NGẮN HẠN` / data$`Nợ ngắn hạn`
data$ROA <- data$`Lợi nhuận sau thuế` / data$`TỔNG TÀI SẢN`
data$FE <- data$`Chi phí tài chính` / data$`TỔNG TÀI SẢN`

names(data)
#THỐNG KÊ MÔ TẢ
install.packages("psych")
library(psych)
#Các biến về chỉ số tài chính
variables <- c("Zmscore", "GDP", "Size", "Current_ratio", "ROA", "FE")
stats_list <- lapply(variables, function(var) {
  describe(data[[var]])
})
stats_df <- do.call(rbind, lapply(1:length(stats_list), function(i) {
  df <- as.data.frame(stats_list[[i]])
  df$Variable <- variables[i]
  df
}))

stats_df <- stats_df[, c(ncol(stats_df), 1:(ncol(stats_df)-1))]
cols_to_keep <- c("Variable", "n", "mean", "sd", "median", "min", "max")
stats_df <- stats_df[, cols_to_keep]

print(stats_df)
write.csv(stats_df, "stats_description.csv", row.names = FALSE)

##Các biến về đặc điểm của CEO
ceo_variables <- c("CEO's Age", "CEO ownership", "CEO's Gender", "CEO's Edu")
ceo_stats_list <- lapply(ceo_variables, function(var) {
  describe(data[[var]])
})
ceo_stats_df <- do.call(rbind, lapply(1:length(ceo_stats_list), function(i) {
  df <- as.data.frame(ceo_stats_list[[i]])
  df$Variable <- ceo_variables[i]
  df
}))
ceo_stats_df <- ceo_stats_df[, c(ncol(ceo_stats_df), 1:(ncol(ceo_stats_df)-1))]
ceo_cols_to_keep <- c("Variable", "n", "mean", "sd", "median", "min", "max")
ceo_stats_df <- ceo_stats_df[, ceo_cols_to_keep]
print(ceo_stats_df)
write.csv(ceo_stats_df, "stats_description_ceo.csv", row.names = FALSE)

install.packages("ggplot2")
library(ggplot2)
data$`CEO's Gender` <- factor(data$`CEO's Gender`, levels = c(0, 1), labels = c("Nam CEO", "Nữ CEO"))
data$`CEO's Edu` <- factor(data$`CEO's Edu`, levels = c(0, 1), labels = c("Cử nhân hoặc khác", "Sau đại học"))

plot_pie_chart_with_labels <- function(data, column_name, title) {
  df <- as.data.frame(table(data[[column_name]]))
  colnames(df) <- c("Category", "Count")
  df$Percentage <- round((df$Count / sum(df$Count)) * 100, 1)
  df$Label <- paste0(df$Percentage, "%")
  ggplot(df, aes(x = "", y = Count, fill = Category)) +
    geom_bar(width = 1, stat = "identity") +
    coord_polar("y", start = 0) +
    labs(title = title, x = NULL, y = NULL) +
    theme_void() +
    theme(legend.title = element_blank()) +
    geom_text(aes(label = Label), position = position_stack(vjust = 0.5)) +
    scale_fill_manual(values = c("Nam CEO" = "#2A629A", "Nữ CEO" = "#FF5F00", 
                                 "Cử nhân hoặc khác" = "#2A629A", "Sau đại học" = "#FF5F00"))
}


plot_pie_chart_with_labels(data, "CEO's Gender", "CEO's Gender")
plot_pie_chart_with_labels(data, "CEO's Edu", "CEO's Education")

# Chia nhóm tuổi
data$AgeGroup <- cut(data$`CEO's Age`,
                     breaks = c(-Inf, 34, 55, Inf),
                     labels = c("Nhóm trẻ tuổi", "Nhóm trung niên", "Nhóm người cao tuổi"))

df <- as.data.frame(table(data$AgeGroup))
colnames(df) <- c("Category", "Count")


df$Percentage <- round((df$Count / sum(df$Count)) * 100, 1)
df$Label <- paste0(df$Percentage, "%")

ggplot(df, aes(x = "", y = Count, fill = Category)) +
  geom_bar(width = 1, stat = "identity") +
  coord_polar("y", start = 0) +
  labs(title = "CEO Age", x = NULL, y = NULL) +
  theme_void() +
  theme(legend.title = element_blank()) +
  geom_text(aes(label = Label), position = position_stack(vjust = 0.5)) +
  scale_fill_manual(values = c("Nhóm trẻ tuổi" = "#FFFAE6", "Nhóm trung niên" = "#2A629A", "Nhóm người cao tuổi" = "#FF5F00"))


#Sở hữu của CEO 
ggplot(data, aes(y = `CEO ownership`)) +
  geom_boxplot(fill = "#2A629A", color = "black") +
  labs(title = "Tỷ lệ sở hữu của CEO", y = "Tỷ lệ sở hữu") +
  theme_minimal()

install.packages("fastDummies")
library(fastDummies)
data$`CEO's Gender` <- ifelse(data$`CEO's Gender` == 'Nam CEO', 0, 1)
data$`CEO's Edu` <- ifelse(data$`CEO's Edu` == 'Sau đại học', 1, 0)

#TRỤC QUAN HÓA DỮ LIỆU
install.packages("gridExtra")
library(ggplot2)
library(gridExtra)
# Biểu đồ phân phối
p1 <- ggplot(data, aes(x = Size)) + 
  geom_histogram(aes(y = ..density..), bins = 30, fill = "blue", alpha = 0.5) +
  geom_density(color = "red") +
  ggtitle("Distribution of Size") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_x_continuous(expand = c(0, 0)) + 
  scale_y_continuous(expand = c(0, 0))

p2 <- ggplot(data, aes(x = ROA)) + 
  geom_histogram(aes(y = ..density..), bins = 30, fill = "blue", alpha = 0.5) +
  geom_density(color = "red") +
  ggtitle("Distribution of ROA") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_x_continuous(expand = c(0, 0)) + 
  scale_y_continuous(expand = c(0, 0))

p3 <- ggplot(data, aes(x = Current_ratio)) + 
  geom_histogram(aes(y = ..density..), bins = 30, fill = "blue", alpha = 0.5) +
  geom_density(color = "red") +
  ggtitle("Distribution of Current_ratio") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_x_continuous(expand = c(0, 0)) + 
  scale_y_continuous(expand = c(0, 0))

p4 <- ggplot(data, aes(x = FE)) + 
  geom_histogram(aes(y = ..density..), bins = 30, fill = "blue", alpha = 0.5) +
  geom_density(color = "red") +
  ggtitle("Distribution of FE") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_x_continuous(expand = c(0, 0)) + 
  scale_y_continuous(expand = c(0, 0))

# Biểu đồ hộp
p5 <- ggplot(data, aes(y = Size)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of Size") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

p6 <- ggplot(data, aes(y = ROA)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of ROA") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

p7 <- ggplot(data, aes(y = Current_ratio)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of Current_ratio") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

p8 <- ggplot(data, aes(y = FE)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of FE") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))


grid.arrange(p1, p2, p3, p4, ncol = 2)
grid.arrange(p5, p6, p7, p8, ncol = 2)


#XỬ LÝ NGOẠI LAI 
install.packages("DescTools")
library(DescTools)
library(ggplot2)
library(gridExtra)

# Xử lý ngoại lai bằng Winsorize
data$Size <- Winsorize(data$Size, probs = c(0.01, 0.99))
data$ROA <- Winsorize(data$ROA, probs = c(0.01, 0.99))
data$Current_ratio <- Winsorize(data$Current_ratio, probs = c(0.01, 0.99))
data$FE <- Winsorize(data$FE, probs = c(0.01, 0.99))

# Vẽ lại biểu đồ hộp sau khi xử lý ngoại lai
p5 <- ggplot(data, aes(y = Size)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of Size (After Winsorizing)") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

p6 <- ggplot(data, aes(y = ROA)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of ROA (After Winsorizing)") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

p7 <- ggplot(data, aes(y = Current_ratio)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of Current_ratio (After Winsorizing)") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

p8 <- ggplot(data, aes(y = FE)) + 
  geom_boxplot(fill = "lightblue") +
  ggtitle("Box Plot of FE (After Winsorizing)") +
  theme_minimal() +
  theme(panel.grid = element_blank(), panel.background = element_blank(), 
        axis.line = element_line(color = "black")) +
  scale_y_continuous(expand = c(0, 0))

grid.arrange(p5, p6, p7, p8, ncol = 2)


#MA TRẬN TƯƠNG QUAN GIỮA CÁC BIẾN 

variables <- c('Zmscore', 'Size', 'ROA', 'Current_ratio', 'FE', 'GDP', 'Lnageceo', 'CEO ownership', "CEO's Gender", "CEO's Edu")
selected_data <- data[, variables]
cor_matrix <- cor(selected_data, use = "complete.obs")  

install.packages("corrplot")
library(corrplot)
corrplot(cor_matrix, method = "circle")
write.csv(cor_matrix, file = "correlation_matrix.csv")

# BIỂU ĐỒ THỂ HIỆN MỨC ĐỘ TƯƠNG QUAN GIỮA MỘT SỐ BIẾN ĐỘC LẬP VỚI BIẾN PHỤ THUỘC
library(ggplot2)
library(gridExtra)

plot1 <- ggplot(data, aes(x = ROA, y = Zmscore)) +
  geom_point() +
  geom_smooth(method = "lm", col = "red") +
  labs(title = "Scatter Plot of Zm_score vs ROA",
       x = "ROA",
       y = "Zmscore")

plot2 <- ggplot(data, aes(x = Current_ratio, y = Zmscore)) +
  geom_point() +
  geom_smooth(method = "lm", col = "red") +
  labs(title = "Scatter Plot of Zm_score vs Current_ratio",
       x = "Current_ratio",
       y = "Zmscore")

plot3 <- ggplot(data, aes(x = FE, y = Zmscore)) +
  geom_point() +
  geom_smooth(method = "lm", col = "red") +
  labs(title = "Scatter Plot of Zm_score vs FE",
       x = "FE",
       y = "Zmscore")

grid.arrange(plot1, plot2, plot3, ncol = 1)

## HỒI QUY 
install.packages("plm")
install.packages("dplyr")
install.packages("data.table")

library(data.table)
library(dplyr)
data <- data %>%
  mutate(
    `CEO ownership` = ifelse(is.na(`CEO ownership`), median(`CEO ownership`, na.rm = TRUE), `CEO ownership`),
    `CEO's Gender` = ifelse(is.na(`CEO's Gender`), median(`CEO's Gender`, na.rm = TRUE), `CEO's Gender`),
    `CEO's Edu` = ifelse(is.na(`CEO's Edu`), median(`CEO's Edu`, na.rm = TRUE), `CEO's Edu`),
    `Lnageceo` = ifelse(is.na(`Lnageceo`), median(`Lnageceo`, na.rm = TRUE), `Lnageceo`)
  )
setnames(data, old = c("CEO ownership", "CEO's Gender", "CEO's Edu"), 
         new = c("CEO_ownership", "CEO_Gender", "CEO_Edu"))
data <- data %>% 
  mutate(id = as.numeric(factor('Mã CK')))


#HỒI QUY MÔ HÌNH FEM 
library(plm)
library(stargazer)
library(car)


pdata <- pdata.frame(data, index = c("id", "Năm"))
print(names(data))

model1 <- plm(Zmscore ~ Size + ROA + Current_ratio + FE + GDP + Lnageceo + CEO_ownership + CEO_Gender + CEO_Edu, 
              data = pdata, 
              model = "within")
summary(model1)

model2 <- plm(Zmscore ~ Size + ROA + Current_ratio + FE + GDP, 
              data = pdata, 
              model = "within")
summary(model2)

rob_se1 <- sqrt(diag(vcovHC(model1, type = "HC1")))
rob_se2 <- sqrt(diag(vcovHC(model2, type = "HC1")))

stargazer(model1, model2,
          digits = 3,
          header = FALSE,
          type = "latex",
          se = list(rob_se1, rob_se2),
          title = "Factors affecting financial distress of enterprises in Vietnam",
          model.numbers = FALSE,
          column.labels = c("(1)", "(2)"))
