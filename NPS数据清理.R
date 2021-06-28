#设置数据路径
ques_path = "问卷数据.xlsx"
user_path = "用户数据.xlsx"

#问卷数据处理
##读取问卷数据
library(readxl)
library(tidyverse)
q1 <- read_excel(ques_path)

##替换-3为缺失值
q1m <- data.frame(q1)
q1m[q1m==-3] <- NA
q1m[q1m=="(空)"] <- NA
q1m[q1m=="NULLTYPE"] <- NA
q2 <- as.tibble(q1m)

##调整NPS值
q3 <- q2 %>% 
  mutate(结合您对课程.服务.教具等各方面的体验.您向亲朋好友推荐叽里呱啦课程的可能性有多大.=结合您对课程.服务.教具等各方面的体验.您向亲朋好友推荐叽里呱啦课程的可能性有多大.- 1,
               
##删去填写时间太短的个案               
           所用时间 = str_replace(所用时间, "秒",""),
           所用时间 = as.numeric(所用时间)) %>% 
  filter(所用时间>=30) %>% 

##以呱号识别，删去重复填写的数据
   distinct(来源详情,.keep_all = T)



#用户数据处理
a <- NA
for(i in 1:length(excel_sheets(user_path))){
  ##读取用户数据
  u1 <- read_excel(user_path, sheet = i, col_names = F) 
  ##合并不同期用户数据
  a <- rbind(a, u1)
}
##整理用户数据
colnames(a)=a[2, ]
u2 <- a[-(1:2),] %>% 
  ##创建期数变量
  mutate("期数"= str_sub(`9块9开课日期（班主任运营期）`,6,10)) %>% 
  mutate(期数=str_replace(期数,"-","")) %>% 
  select(`用户呱号`, `孩子年龄`, `等级城市`, `第一周结束后9块9体验课完课数`, `正价课付费金额`, `9块9课程版本`,期数) %>% 
  ##添加目标变量
  mutate("是否完课" = if_else(`第一周结束后9块9体验课完课数` == "5", "已完课","未完课"),
     正价课付费金额 = as.numeric(正价课付费金额)) %>% 
  mutate("课包类型"=case_when(正价课付费金额>2900 ~ "大课包",
                                 正价课付费金额<=2900& 正价课付费金额>0 ~ "小课包",
                                 TRUE ~ "未付费"),
         `用户呱号`= as.numeric(`用户呱号`)) %>% 
  select(-(4:5)) %>% 
  filter(课包类型!="小课包")

#合并问卷数据和用户数据
d1 <- q3 %>% 
  inner_join(u2, by = c("来源详情"="用户呱号")) %>% 
  distinct(来源详情,.keep_all = T) %>% 
  ##合并小类别
  mutate(孩子年龄= recode(孩子年龄,"0-1岁"="0-2岁","1-2岁" = "0-2岁",
                             "7-8岁"="7岁及以上","8岁以上"= "7岁及以上"),
             等级城市= recode(等级城市,"4线"="4线及以下","5线"= "4线及以下"))


#计算权重
##问卷数据权重
p1 <- d1 %>% group_by(课包类型) %>% 
  count() %>% 
  mutate("p1"=n/count(d1))
##总体数据权重
p2 <- u2 %>% 
  group_by(课包类型) %>% 
  count() %>% 
  mutate("p2"=n/count(u2)) %>% 
  left_join(p1,by=c("课包类型")) %>% 
  mutate("权重"=p2/p1) 

#将权重合并入匹配数据
weight <- p2 %>% 
  select(1,6)

d2 <- d1 %>% 
  left_join(weight) %>% 
  mutate(权重=as.vector(unlist(权重))) %>% 
  
  ##添加NPS标签并给开放题排序
  mutate(NPS=case_when(结合您对课程.服务.教具等各方面的体验.您向亲朋好友推荐叽里呱啦课程的可能性有多大.>=8~ "推荐者",
                             结合您对课程.服务.教具等各方面的体验.您向亲朋好友推荐叽里呱啦课程的可能性有多大.<=6~ "贬损者",
                             TRUE~"中立者")) %>% 
  mutate(NPS=factor(NPS,levels = c("推荐者","贬损者","中立者"))) %>% 
  arrange(NPS,请问您比较愿意推荐叽里呱啦课程的原因是.,请问您不太愿意推荐叽里呱啦课程的原因是.,请问您对叽里呱啦课程哪些方面比较满意.哪些方面不太满意.)
  

#导出Excel文件
library(xlsx)
write.xlsx(d2,"整合数据.xlsx","整合数据")

