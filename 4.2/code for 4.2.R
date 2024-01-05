library(network)
library(btergm)
setwd("D:/量化数据以及大视频/科研/贸易网路图/第一篇论文/论文最后提交的数据和代码【可以复现所有结果】/第二部分")


load("net_upstream.RData")
load("COLONIZIAL_RELATIONSHIP.RData")
load("LINGUISTIC_PROXIMITY.RData")
load("CAPTITAL_DISTANCE.RData")


# Regression
model1<-btergm(net_upstream ~ edges 
               +nodeocov("lnGDP")+nodeocov("lnGDP_CAP")+nodeocov("CO2_EMISSION_PER")
               +nodeicov("lnGDP")+nodeicov("lnGDP_CAP")+nodeicov("CO2_EMISSION_PER")
               +nodematch("LANDLOCKED",diff = FALSE)
               +edgecov(LINGUISTIC_PROXIMITY)
               +edgecov(COLONIZIAL_RELATIONSHIP)
               +edgecov(CAPTITAL_DISTANCE), 
               R =1000, parallel = "snow", ncpus = 16)
model_op1<-btergm.se(model1, print = TRUE)

model2<-btergm(net_upstream~ edges 
               +nodeocov("lnGDP")+nodeocov("lnGDP_CAP")+nodeocov("CO2_EMISSION_PER")
               +nodeicov("lnGDP")+nodeicov("lnGDP_CAP")+nodeicov("CO2_EMISSION_PER")
               +nodematch("LANDLOCKED",diff = FALSE)
               +edgecov(LINGUISTIC_PROXIMITY)
               +edgecov(COLONIZIAL_RELATIONSHIP)
               +edgecov(CAPTITAL_DISTANCE)
               +mutual, 
               R =1000, parallel = "snow", ncpus = 16)
model_op2<-btergm.se(model2, print = TRUE)


model3<-btergm(net_upstream~ edges +nodeocov("lnGDP")+nodeocov("lnGDP_CAP")+nodeocov("CO2_EMISSION_PER")
               +nodeicov("lnGDP")+nodeicov("lnGDP_CAP")+nodeicov("CO2_EMISSION_PER")
               +nodematch("LANDLOCKED",diff = FALSE)
               +edgecov(LINGUISTIC_PROXIMITY)
               +edgecov(COLONIZIAL_RELATIONSHIP)
               +edgecov(CAPTITAL_DISTANCE)
               +mutual
               +gwidegree(0.1,T) 
               +gwesp(0.1,T)
               +gwdsp(0.1,T), 
               R =1000, parallel = "snow", ncpus = 16)
model_op3<-btergm.se(model3, print = TRUE)


model4<-btergm(net_upstream~ edges +nodeocov("lnGDP")+nodeocov("lnGDP_CAP")+nodeocov("CO2_EMISSION_PER")
               +nodeicov("lnGDP")+nodeicov("lnGDP_CAP")+nodeicov("CO2_EMISSION_PER")
               +nodematch("LANDLOCKED",diff = FALSE)
               +edgecov(LINGUISTIC_PROXIMITY)
               +edgecov(COLONIZIAL_RELATIONSHIP)
               +edgecov(CAPTITAL_DISTANCE)
               +mutual
               +gwidegree(0.1,T) 
               +gwesp(0.1,T)
               +gwdsp(0.1,T)
               +memory(type = "stability")
               +delrecip, 
               R =1000, parallel = "snow", ncpus = 16)
model_op4<-btergm.se(model4, print = TRUE)