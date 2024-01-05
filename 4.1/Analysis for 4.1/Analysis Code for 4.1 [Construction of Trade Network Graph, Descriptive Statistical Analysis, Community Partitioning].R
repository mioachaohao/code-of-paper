
environment_file_path="D:\code"

setwd(environment_file_path)
#1导入包，数据
# 定义需要检查的包列表
packages <- c("tidyverse", "igraph", "statnet", "sysfonts", 
              "showtext", "Cairo", "readxl", "maps","openxlsx")

# 循环遍历包列表
for (pkg in packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg)
  }
}


load("Data for the Upstream, Midstream, and Downstream of the Photovoltaic Industry.RData")
country_code <- read_excel("country_code.xlsx",sheet="English")
country_code<-select(country_code,i=country_code,name=iso_3digit_alpha)

Descriptive_feature=list()
Descriptive_feature$years<-c(1995:2019)
upstream_max_10_degree_countries<-list()
upstream_core_commulity<-list()

for(a in 1995:2019)
{b=a-1994
for( i in c("upstream")){
  owd<-paste(a,paste("_",i),sep ="")
  data_export<-get(paste("data_",a,paste("_",i,sep=""),sep=""))%>%group_by(i)%>%mutate(rate=v/sum(v))
  data_import<-get(paste("data_",a,paste("_",i,sep=""),sep=""))%>%group_by(j)%>%mutate(rate=v/sum(v))
  export<-data_export%>%group_by(i)%>%summarise(export=ifelse(rate==1,1,((sum(rate^2))^0.5-(1/n())^0.5)/(1-(1/n())^0.5)))
  import<-data_import%>%group_by(j)%>%summarise(import=ifelse(rate==1,1,((sum(rate^2))^0.5-(1/n())^0.5)/(1-(1/n())^0.5)))
  export<-unique(export)
  import<-unique(import)
  data_export<-left_join(data_export,export,by="i")
  data_export<-left_join(data_export,import,by="j")
  data_import<-left_join(data_import,export,by="i")
  data_import<-left_join(data_import,import,by="j")
  #计算所有出口国家的 出口依赖
  extd<-mutate(data_export,extd=(rate^2*export/import))
  #计算所有进口国家 进口依赖
  imtd<-mutate(data_import,imtd=(rate^2*import/export))
  extd$extd[is.infinite(extd$extd)]<-18
  
  imtd$imtd[is.infinite(imtd$imtd)]<-1
  extd<-select(extd,i,j,extd)
  imtd<-select(imtd,i,j,imtd)
  td<-full_join(extd,imtd,by=c("i","j"))
  
  #去掉td中的异常值
  td<-na.omit(td)
  td$td<-td$extd+td$imtd
  td$td=(log(td$td)-min(log(td$td)))/(max(log(td$td))-min(log(td$td)))+0.1
  td<-data.frame(i=td$i,j=td$j,weight=td$td)
  td<-td%>%group_by(i)%>%top_n(n=5,wt=weight)%>%ungroup()
  td<-left_join(td,country_code,by="i")
  td<-left_join(td,country_code,by=c("j"="i"))
  td<-select(td,i=name.x,j=name.y,weight)
  #去掉没有对应名字的编号
  td<-na.omit(td)
  g<-graph.data.frame(td[1:2],directed = TRUE)
  V(g)$degree<-igraph::degree(g,mode = "in")
  E(g)$weight<-td$weight
}
g_upstream<-g
upstream_vertices_names<-as.vector(V(g_upstream)$name)


Descriptive_feature$upstream_nodes<-append(Descriptive_feature$upstream_nodes,vcount(g_upstream))
Descriptive_feature$upstream_edges<-append(Descriptive_feature$upstream_edges,ecount(g_upstream))
Descriptive_feature$upstream_density<-append(Descriptive_feature$upstream_density,graph.density(g_upstream))
Descriptive_feature$upstream_reciprocity<-append(Descriptive_feature$upstream_reciprocity,reciprocity(g_upstream))
Descriptive_feature$upstream_transitivity<-append(Descriptive_feature$upstream_transitivity,transitivity(g_upstream))
Descriptive_feature$upstream_average_length<-append(Descriptive_feature$upstream_average_length,average.path.length(g_upstream))
Descriptive_feature$upstream_diamter<-append(Descriptive_feature$upstream_diamter,diameter(g_upstream))



###随机模拟


# 获取上游网络的节点数和边数
num_nodes <- vcount(g_upstream)
num_edges <- ecount(g_upstream)

# 计算边的概率
p <- 2 * num_edges / (num_nodes * (num_nodes - 1))

# 生成随机网络
g_random <- erdos.renyi.game(n = num_nodes, p = p, directed = TRUE)

# 计算随机网络的描述性统计

Descriptive_feature$random_transitivity <- append(Descriptive_feature$random_transitivity, transitivity(g_random))
Descriptive_feature$random_reciprocity <- append(Descriptive_feature$random_reciprocity, reciprocity(g_random))
Descriptive_feature$random_average_length <- append(Descriptive_feature$random_average_length, average.path.length(g_random))
Descriptive_feature$random_diameter <- append(Descriptive_feature$random_diameter, diameter(g_random))


}

# 请注意，将反斜杠转义为双反斜杠（\\）
write.csv(Descriptive_feature, "第一部分结果/Descriptive_feature.csv")


wb_nodelist <- createWorkbook()
wb_edgelist <- createWorkbook()
for(a in 1995:2019)
{b=a-1994
for( i in c("upstream")){
  owd<-paste(a,paste("_",i),sep ="")
  data_export<-get(paste("data_",a,paste("_",i,sep=""),sep=""))%>%group_by(i)%>%mutate(rate=v/sum(v))
  data_import<-get(paste("data_",a,paste("_",i,sep=""),sep=""))%>%group_by(j)%>%mutate(rate=v/sum(v))
  export<-data_export%>%group_by(i)%>%summarise(export=ifelse(rate==1,1,((sum(rate^2))^0.5-(1/n())^0.5)/(1-(1/n())^0.5)))
  import<-data_import%>%group_by(j)%>%summarise(import=ifelse(rate==1,1,((sum(rate^2))^0.5-(1/n())^0.5)/(1-(1/n())^0.5)))
  export<-unique(export)
  import<-unique(import)
  data_export<-left_join(data_export,export,by="i")
  data_export<-left_join(data_export,import,by="j")
  data_import<-left_join(data_import,export,by="i")
  data_import<-left_join(data_import,import,by="j")
  #计算所有出口国家的 出口依赖
  extd<-mutate(data_export,extd=(rate^2*export/import))
  #计算所有进口国家 进口依赖
  imtd<-mutate(data_import,imtd=(rate^2*import/export))
  extd$extd[is.infinite(extd$extd)]<-18
  
  imtd$imtd[is.infinite(imtd$imtd)]<-1
  extd<-select(extd,i,j,extd)
  imtd<-select(imtd,i,j,imtd)
  td<-full_join(extd,imtd,by=c("i","j"))
  
  #去掉td中的异常值
  td<-na.omit(td)
  td$td<-td$extd+td$imtd
  td$td=(log(td$td)-min(log(td$td)))/(max(log(td$td))-min(log(td$td)))+0.1
  td<-data.frame(i=td$i,j=td$j,weight=td$td)
  td<-td%>%group_by(i)%>%top_n(n=5,wt=weight)%>%ungroup()
  td<-left_join(td,country_code,by="i")
  td<-left_join(td,country_code,by=c("j"="i"))
  td<-select(td,i=name.x,j=name.y,weight)
  #去掉没有对应名字的编号
  td<-na.omit(td)
  g<-graph.data.frame(td[1:2],directed = TRUE)
  V(g)$degree<-igraph::degree(g,mode = "in")
  E(g)$weight<-td$weight
}
g_upstream<-g
upstream_vertices_names<-as.vector(V(g_upstream)$name)

library (plyr)
cw_upstream<- cluster_louvain(as.undirected(g_upstream,mode ="collapse", edge.attr.comb = igraph_opt("edge.attr.sum")), weights = E(g_upstream)$weight,resolution=0.4)
cw_upstream <- cw_upstream[names(sort(sizes(cw_upstream), decreasing= TRUE))]
cw_upstream <- ldply (cw_upstream, data.frame)
cw_upstream=left_join(as.data.frame(upstream_vertices_names),cw_upstream,by=c("upstream_vertices_names"="X..i.."))
cw_upstream$.id[is.na(cw_upstream$.id)]=0
V(g_upstream)$community<-cw_upstream$.id
V(g_upstream)$degree<-igraph::degree(g_upstream,mode = "in") 
detach("package:plyr")

assign(paste("g_upstream_",a,sep=""),g_upstream)

edgelist_upstream<- data.frame(get.edgelist(g_upstream),E(g_upstream)$weight)
colnames(edgelist_upstream)<-c("Source","Target","weight")
nodelist_upstream<-data.frame( Id=V(g_upstream)$name,label=V(g_upstream)$name,degree=V(g_upstream)$degree,community=V(g_upstream)$community)
assign(paste("nodelist_upstream_",a,sep=""),nodelist_upstream,)


nodelist_sheet_name <- paste("nodelist_", a, sep = "")
addWorksheet(wb_nodelist, sheetName = nodelist_sheet_name)
writeData(wb_nodelist, sheet = nodelist_sheet_name, x = nodelist_upstream)

edgelist_sheet_name <- paste("edgelist_", a, sep = "")
addWorksheet(wb_edgelist, sheetName = edgelist_sheet_name)
writeData(wb_edgelist, sheet = edgelist_sheet_name, x = edgelist_upstream)
}

# 保存Excel文件，覆盖现有文件
saveWorkbook(wb_nodelist, file =  "第一部分结果/wb_nodelist.xlsx", overwrite = TRUE)
saveWorkbook(wb_edgelist, file =  "第一部分结果/wb_edgelist.xlsx", overwrite = TRUE)




