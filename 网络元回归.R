library(netmeta)     
library(meta)        
library(readxl)
library(metafor)
library(rjags)
library(gemtc)

# 获取这个文件里所有的工作表名称
path <- "C:\\Users\\16124\\OneDrive\\桌面\\netmeta模拟数据\\point.xlsx" 
sheet_names <- excel_sheets(path)
print(sheet_names)

# 读取主表
data <- read_excel("C:\\Users\\16124\\OneDrive\\桌面\\netmeta模拟数据\\point.xlsx", sheet = "all")

# 检查列名，确认是否有回归需要的协变量
colnames(data)

# 确保核心列是数值型（非常重要，否则 mtc.network 会报错）
data$responders <- as.numeric(data$responders)
data$sampleSize <- as.numeric(data$sampleSize)

# 构建网络对象
# gemtc 默认识别 responders 和 sampleSize 这两个列名
# 把Excel 表格转换成 R 能够理解的“网络拓扑结构”——并非建模
network_mtc <- mtc.network(data)

# 绘制网络图（修正之前的 vertex.label.cex 问题）
plot(network_mtc,
     use.description = TRUE, vertex.label.cex = 1,
     vertex.size = 5, vertex.shape = "circle", vertex.label.color = "black",
     vertex.label.dist = 1, vertex.label.degree = -pi/2, vertex.label.cex = 200,
     vertex.color = "orange",
     dynamic.edge.width = TRUE,
     edge.color = "darkgrey",
     vertex.label.font = 2)

# Bayesian modeling and regression analysis
# 假设网络是“一致”的，即直接证据和间接证据没有显著冲突
# n.chain = 4：启动 4 条独立的模拟链
# 
model <- mtc.model(network_mtc, 
                   type = "consistency",
                   n.chain = 4, 
                   likelihood = "binom", #二项分布似然函数（Binomial Likelihood）——处理的数据是二分类数据
                   link = "log", 
                   linearModel = "random") #使用随机效应模型，承认不同研究之间存在异质性

# 需要下载JAGS软件
results <- mtc.run(model,
                   n.adapt = 20000, # 预热，保证模拟趋于稳定
                   n.iter = 50000,  # JAGS 会记录下这 50,000 次模拟中的每一个数值（正式实验期）
                   thin = 1) #抽稀间隔，设置为 1 意味着“保留每一次模拟的结果”

summary(results)
forest(relative.effect(results, "Standard_Care"))

# Rank probability plot
rank <- rank.probability(results,
                         preferredDirection = -1) #preferredDirection = -1代表数据越小越好
# beside = TRUE：将不同排名的概率并排堆叠
plot(rank, beside = TRUE)
re_results <- relative.effect(results, t1 = "Standard_Care")
forest(re_results, use.description = TRUE)

# Analysis for other variables (sex, age, smoking status, year, etc.)
studies <- data

# 标准化变量格式
studies$sm <- as.numeric(studies$sm)
studies$Sex <- as.numeric(studies$Sex)

#清除空白量（average_age）
studies <- studies[!is.na(studies$Average_age), ]

#转化变量为二进制变量
studies$`2010` <- ifelse(tolower(studies$`2010`) == "yes", 
                         1, ifelse(tolower(studies$`2010`) == "no", 0, NA))

studies$`2015` <- ifelse(tolower(studies$`2015`) == "yes", 
                         1, ifelse(tolower(studies$`2015`) == "no", 0, NA))

studies$bioreport <- ifelse(tolower(studies$bioreport) == "yes", 
                            1, ifelse(tolower(studies$bioreport) == "no", 0, NA))

studies$drug <- ifelse(tolower(studies$drug) == "yes", 1, ifelse(tolower(studies$drug) == "no", 0, NA))

studies$money <- ifelse(tolower(studies$money) == "yes", 1, ifelse(tolower(studies$money) == "no", 0, NA))


# mtc.network 的唯一作用是创建一个 mtc.network 类的对象。
# 这个对象就像一个“集装箱”，把零散的原始数据（谁比了谁、人数、事件数）和研究背景（年龄、年份等）打包在一起，
# 以便后续的模型计算。
# data.ab()——会自动选择列，其中必须有的核心列包括：study、treatment——在创建文件的时候必须有这两个列名
# 同时data.ab()会有预设——由“似然函数 (Likelihood)”决定的数据列——gemtc列命名非常死板
network <- mtc.network(data.ab = studies,  # 这是构建网状元分析的核心数据()
                       studies = studies)
# studies：它会把剩下的那些它“不认识”的列（如年龄、性别、年份等）存放在 studies 槽位里，留作后续的回归分析使用

# 对纳入的协变量分别进行回归计算
# 将平均年龄减去所有研究的均值
studies$Average_age_centered <- studies$Average_age - mean(studies$Average_age, na.rm = TRUE)

# 重新构建包含中心化变量的网络对象
network <- mtc.network(data = studies, studies = studies)
# 以“平均年龄”作为协变量运行网状元回归
model_age_unrelated <- mtc.model(
  network,
  type = "regression",
  n.chain = 4,
  regressor = list(
    coefficient = "unrelated",      # 每个干预措施拥有独立的回归系数
    variable    = "Average_age_centered", 
    control     = "Standard_Care"   # 请确保这里是你数据集里的对照组名称
  )
)

# 运行模型
results_age <- mtc.run(model_age_unrelated, 
                       n.adapt = 20000, 
                       n.iter = 50000,
                       thin = 1)

# 计算 PSRF (Gelman-Rubin 诊断)——确定模型是否稳定，大于1.1属于不稳定，需要增加迭代次数
gelman.diag(results_age)
plot(results_age)
summary(results_age)

# Model for 'Average Age'——同一个干预措施对于结果的影响使用同一个回归系数——“shared”
model_age_shared <- mtc.model(network, type = "regression", 
                              n.chain = 4,
                              regressor = list(coefficient = 'shared', 
                                                variable = 'Average_age', 
                                                control = 'Standard_Care'))
results_age_shared <- mtc.run(model_age, n.adapt = 5000, n.iter = 20000, thin = 1)

gelman.diag(results_age_shared)
plot(results_age_shared)

summary(results_age)

