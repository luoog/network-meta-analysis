library(netmeta)     # For frequentist network meta-analysis
library(meta)        # For conventional meta-analysis
library(readxl)
library(metafor)

data_all <- read_excel("C:\\Users\\16124\\OneDrive\\桌面\\netmeta模拟数据\\point.xlsx", 
                       sheet = "all")

data_all$responders <- as.numeric(data_all$responders)
data_all$sampleSize <- as.numeric(data_all$sampleSize)

p1 <- pairwise(treatment, event = responders, n = sampleSize,
               studlab = study, data = data_all, sm = "RR")

netconnection <- netconnection(p1)
print(netconnection)

#创建网络（频率主义网络图）
net1 <- netmeta(TE, seTE, treat1, treat2, studlab, data = p1, sm = "RR",
                comb.fixed = FALSE, comb.random = TRUE, reference = "Standard_Care")
names(net1)

# 优化后的网状图绘图代码
study_counts <- net1$k.trts
print(study_counts)

netgraph(net1)
netgraph(net1, plastic = FALSE)
netgraph(net1,
         iterate = FALSE,
         plastic = FALSE,               # 禁用立体线条，使图面整洁
         thickness = "number.of.studies", # 线条粗细映射研究数量
         lwd = 2,                       # 基础线宽
         points = TRUE,                 # 显示节点圆点
         col.points = "dodgerblue",     # 节点颜色
         col = "gray70",                # 使用浅灰色连线
         cex.points = 5 * sqrt(study_counts / max(study_counts)),
         cex = 0.7,                     # 缩小标签字体
         offset = 0.015,                # 调整标签偏移量
         number.of.studies = FALSE,
         alpha.transparency = 0.4       # 增加透明度，减轻视觉压力
)

# 绘制森林图：以 Standard care 为参照
# sortvar = TE 表示按效应量大小排序，让图表更清晰
forest(net1, 
       reference.group = "Standard_care", 
       sortvar = TE,
       xlim = c(0.1, 40),     # 根据你的数据范围调整坐标轴
       smlab = "Relative Risk (95% CI)",
       label.left = "Favours Standard care", 
       label.right = "Favours Intervention")

# 查看随机效应模型的表格结果
netleague <- netleague(net1, bracket = "(", digits = 2)
netleague_table <- netleague$random
write.csv(netleague_table, "C:\\Users\\16124\\OneDrive\\桌面\\netmeta模拟数据\\netleague.csv")
#R控制台显示不齐全，会切割显示，CSV文件可以使用Excel打开

# 生成治疗排名（P得分，越小越认为是“bad”）
# P-score 绝对不是 P 值（p-value）
netrank(net1, small.values = "bad")

# 绘制森林图，以 Standard_Care 为对照组
# sortvar = -Pscore 表示按排名从高到低排序
forest(net1, 
       reference.group = "Standard_Care", 
       sortvar = -Pscore, 
       smlab = "Relative Risk (95% CI)", 
       drop.reference.group = TRUE, # 不在图中显示对照组自己比自己
       label.left = "Favors Standard Care", 
       label.right = "Favors Interventions",
       col.square = "dodgerblue")

# Generate heatmap to assess inconsistency
netheat(net1, random = TRUE)

# Display network consistency check
# 节点拆分法（Node-splitting）
netsplit(net1)
sink("C:\\Users\\16124\\OneDrive\\桌面\\netmeta模拟数据\\output.txt")          # Redirect output to a file
print(netsplit(net1))       # Save network split details
sink()

# Publication Bias Analysis using Funnel Plot
# rainbow() 会生成一组彩虹渐变色。
colors <- rainbow(45)      
# pch = point character
# R中可用的形状较少，只有26种
# 循环重复这些形状 直到变成 45 个
pch_values <- rep(0:25, length.out = 45)  # Set point character values

# Create funnel plot with treatment order
funnel(net1,
       order = c("Placebo", "Standard_Care", "Face_to_Face", "Phone", "SMS",
                 "Email", "Web", "App", "Multicomponent_Intervention", "Interactive_SMS",
                 "Interactive_Web", "Interactive_app", "Customized_SMS", "Customized_Email",
                 "Customized_Web", "Customized_App", "Group_Customized_SMS",
                 "Group_Customized_App", "Group_Customized_Web", "Group_Customized_Phone",
                 "Group_Customized_Multicomponent_Intervention"),
       pch = pch_values,
       col = colors[1:45],  # Use 45 unique colors
       linreg = TRUE,       # linreg = TRUE (线性回归线)
       legend = FALSE)

# 根据研究的具体需求，对特定的干预措施对比进行更深入的漏斗图分析
# 1. 从原始对比数据 p1 中筛选出特定对比的研究
sub_data_1 <- p1[(p1$treat1 == "App" & p1$treat2 == "Standard_Care") | 
                  (p1$treat1 == "Standard_Care" & p1$treat2 == "App"), ]
print(sub_data_1)

# 2. 标准 funnel 函数（两两对比模式）
# 两两对比是用的是meta包中的funnel函数，建议使用两个包，指定清楚
# 之前全局的funnel函数用的是netmeta中的函数
m_sub <- metagen(TE, seTE, data = sub_data_1, sm = "RR")
meta::funnel(m_sub)
#通常建议在研究数量大于10时才执行 Egger's 检验。
# 使用 as.rma() 将 meta 对象转换为 metafor 兼容对象
# yi 对应对数效应量 (TE)，sei 对应标准误 (seTE)
#使用 REML（受限最大似然法)
m_rma_final <- metafor::rma(yi = TE, sei = seTE, data = sub_data_1, method = "REML")
egg_test <- regtest(m_rma_final)

# 再次尝试打印结果
print(egg_test)



