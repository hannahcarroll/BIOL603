library(readxl)
library(MuMIn)
library(parallel)
library(dplyr)
library(performance)
library(effectsize)
library(factoextra)
library(flextable)
library(openxlsx)
library(cowplot)

# install.packages("remotes")
 remotes::install_github("Sebastien-Le/YesSiR")
library(YesSiR)

##############
# Output model tables to xlsx using flextable
set_flextable_defaults(
  font.size = 10, theme_fun = theme_vanilla, font.family = "arial",
  padding = 6,
  background.color = "#FFFFFF")

wb <- createWorkbook()

#############

#forams <- read_excel("foram cal update 10_24 with atlas temps.xlsx", sheet = "Sheet1")
forams <- read_excel("foram cal update 2_25_23 atlas temps_per-pia-menupdates.xlsx", sheet = "Sheet1")
# Don't want to scale the response
forams2 <- forams %>% select(-D47ICDES) %>% mutate(across(where(is.numeric), scale)) %>% mutate(D47ICDES = forams$D47ICDES)

forams2$BenthicPlanktic <- ifelse(grepl("Benthic", forams2$Habitat), "Benthic", 
                               ifelse(forams2$Habitat == "Bulk", NA, "Planktic"))

forams2$BenthicPlankticcoded <- ifelse(forams2$BenthicPlanktic == "Benthic", 0,
                                       ifelse(forams2$BenthicPlanktic == "Planktic", 1, NA))

forams2$Symbiontscoded <- ifelse(forams2$Symbionts == "No Photosymbionts", 0,
                                 ifelse(forams2$Symbionts == "Photosymbionts", 1, NA))

forams$BenthicPlanktic <- ifelse(grepl("Benthic", forams$Habitat), "Benthic", 
                                  ifelse(forams$Habitat == "Bulk", NA, "Planktic"))

forams$BenthicPlankticcoded <- ifelse(forams$BenthicPlanktic == "Benthic", 0,
                                       ifelse(forams$BenthicPlanktic == "Planktic", 1, NA))

forams$Symbiontscoded <- ifelse(forams$Symbionts == "No Photosymbionts", 0,
                                 ifelse(forams$Symbionts == "Photosymbionts", 1, NA))
  
summarytable <- forams %>%
  group_by(Dataset, Region) %>%
  summarize(n = n())

forams3 <- forams2[!is.na(forams2$BenthicPlanktic),]
forams4 <- forams[!is.na(forams$BenthicPlanktic),]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "forams3")

#fullmodel <- lm(calcTiso ~ D47ICDES*Dataset*Ocean*region*Depth*Genus*Habitat*carbsat*Symbionts, data = forams)
#fullmodel <- lm(calcTiso ~ D47ICDES+Dataset+region+Depth+Genus+BenthicPlanktic+carbsat+Symbionts+Salinity, data = forams2)
#summary(fullmodel)



fullmodel2 <- lm(D47ICDES ~ calcTiso*Dataset*carbsat*BenthicPlanktic+Salinity+Depth+Symbionts+Ocean, data = forams3, na.action = "na.fail")
fullmodel2 <- lm(D47ICDES ~ calcTiso+Dataset*Depth*Symbionts*Salinity+carbsat*BenthicPlanktic+Ocean, data = forams3, na.action = "na.fail")
summary(fullmodel2)


names(forams3)
fullmodel2.dredge <- dredge(fullmodel2, cluster = clust, fixed = "calcTiso")
plot(subset(fullmodel2.dredge, delta < 4 | df == min(df)))

# reducedmodel <- lm(D47ICDES ~ calcTiso+Dataset*carbsat+Symbionts, data = forams3, na.action = "na.fail") 10/24 dataset update
# reducedmodel2 <- lm(D47ICDES ~ calcTiso+Dataset+carbsat+Symbionts, data = forams3, na.action = "na.fail") 10/24 dataset update
#reducedmodel3 <- lm(D47ICDES ~ calcTiso+Dataset*carbsat+Symbionts, data = forams4, na.action = "na.fail")

# Final model 3/3/23
reducedmodel <- lm(D47ICDES ~  calcTiso + BenthicPlanktic*carbsat + Dataset*Depth + Dataset*Symbionts, data = forams3, na.action = "na.fail")
summary(reducedmodel)

reducedmodelunscaled <- lm(D47ICDES ~  calcTiso + BenthicPlanktic*carbsat + Dataset*Depth + Dataset*Symbionts, data = forams4, na.action = "na.fail")
summary(reducedmodelunscaled)

reducedmodelcoded <- lm(D47ICDES ~  calcTiso + BenthicPlankticcoded*carbsat + Dataset*Depth + Dataset*Symbiontscoded, data = forams3, na.action = "na.fail")
summary(reducedmodelcoded)

library(officer)

tf <- tempfile(fileext = ".docx")

reducedmodelforams3 <- as_flextable(reducedmodel)




# For collinearity check # OK #
reducedmodel2 <- lm(D47ICDES ~  calcTiso + BenthicPlanktic + carbsat + Dataset + Depth + Symbionts, data = forams3, na.action = "na.fail")

anova(reducedmodel, reducedmodel2)
AIC(reducedmodel2)
reducedmodel
summary(reducedmodel)
summary(reducedmodel2)
plot(reducedmodel)
car::Anova(reducedmodel, type = 2)
car::Anova(reducedmodel3, type = 2)
check_collinearity(reducedmodel2)
check_distribution(reducedmodel2)
reducedmodelt2 <- car::Anova(reducedmodel, type = 2)
eta_squared(car::Anova(reducedmodel, type = 2))


############################# Get back-transformed model coefficients if Y is scaled:

# Modified from Ben Bolker's stack overflow post here:
# https://stackoverflow.com/questions/23642111/how-to-unscale-the-coefficients-from-an-lmer-model-fitted-with-a-scaled-respon


rescale.coefs <- function(beta,mu,sigma) {
  beta2 <- beta ## inherit names etc.
  beta2[-1] <- sigma[1]*beta[-1]/sigma[-1]
  beta2[1]  <- sigma[1]*beta[1]+mu[1]-sum(beta2[-1]*mu[-1])
  beta2
}

# reducedmodel is the final model for that particular subset of z-scored data.
# my forams4 is the full, unscaled dataset with any rows dropped where symbiont information is missing. Replace with the appropriate unscaled dataset for whichever model you're working on.
b1S <- coef(reducedmodel)
summary(m)$coefficients
sterr <- summary(reducedmodel)$coefficients[,2]
mo <- forams4 %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))
  
icol <- which(colnames(forams4)=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(forams4, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(forams4, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

# If unscaled Y:

# A = As - Bs*Xmean/sdx
# B = Bs/sdx  Bs/s*X = Bs*(X/s)
# thus the regression is,
# 
# Y = As - Bs*Xmean/sdx + Bs/sdx * X
# where
# 
# As = intercept from the scaled regression
# Bs = slope from the scaled regression
# Xmean = the mean of the scaled predictor variable
# sdx = the standard deviation of the predictor variable

options(es.use_symbols = TRUE)
eta_squared(reducedmodel, partial = TRUE)

# Turn off sequential sums of squares
eta_squared(car::Anova(reducedmodel, type = 2))

names(forams)

fullmodel.check <- lm(D47ICDES ~ calcTiso+Dataset+carbsat+BenthicPlanktic, data = forams3)

check_collinearity(fullmodel.check)
check_distribution(fullmodel.check)
check_heteroskedasticity(fullmodel.check)
check_model(fullmodel.check)

# PCA for reduced model

# reducedmodel <- lm(D47ICDES ~ calcTiso+Dataset+carbsat*BenthicPlanktic, data = forams3, na.action = "na.fail")

# Prepare dataset

forampca <- forams3[ , c("calcTiso", "Dataset", "carbsat", "BenthicPlanktic", "Depth", "Salinity")]
names(forampca) <- c("Temperature","Dataset","Carbonate saturation","BenthicPlanktic","Depth","Salinity") 

f.pca <- prcomp(forampca[,c(1,3,5,6)], scale = TRUE)

fviz_eig(f.pca)

fviz_pca_var(f.pca,
             col.var = "contrib", # Color by contributions to the PC
             gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"),
             repel = TRUE     # Avoid text overlapping
) + labs(color = "Contribution", main = NULL) +
  theme_bw()
ggsave("fullmodelpca.png", dpi = 900, height = 5, width = 5, units = "in", scale = 1.5)

fviz_pca_ind(f.pca,
             col.ind = "cos2", # Color by the quality of representation
             gradient.cols = c("#00AFBB", "#E7B800", "#FC4E07"),
             repel = TRUE     # Avoid text overlapping
)


dim1 <- fviz_contrib(f.pca, "var") + ggtitle(NULL) + ylim(0,60) + ylab(NULL)
dim2 <- fviz_contrib(f.pca, "var", axes = 2) + ggtitle(NULL) + ylim(0,60) + ylab(NULL)

plot_grid(dim1, dim2, labels = c("Dimension 1", "Dimension 2")) + draw_label("Contributions (%)", x=0, y=0.5, vjust= 0.5,angle=90)

summary(f.pca)

# Benthic vs planktic

benthonly <- forams3[forams3$BenthicPlanktic == "Benthic",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "benthonly")


fullbenthmodel <- lm(D47ICDES ~ calcTiso*Dataset*Depth*Salinity+carbsat+Ocean*Habitat, data = benthonly, na.action = "na.fail")
summary(fullbenthmodel)

fullbenthmodel.dredge <- dredge(fullbenthmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullbenthmodel.dredge, delta < 4 | df == min(df)))

reducedbenthmodel <- lm(D47ICDES ~ calcTiso+carbsat+Salinity, data = benthonly, na.action = "na.fail")
reducedbenthmodel2 <- lm(D47ICDES ~ calcTiso+carbsat+Salinity+Habitat, data = benthonly[benthonly$Habitat != "Benthic",], na.action = "na.fail")
anova(reducedbenthmodel, reducedbenthmodel2)
summary(reducedbenthmodel2) # 2 is worse
eta_squared(car::Anova(reducedbenthmodel, type = 2))

plankonly <- forams3[forams3$BenthicPlanktic == "Planktic",]

clusterExport(clust, "plankonly")

fullplankmodel <- lm(D47ICDES ~ calcTiso*Dataset*carbsat*Symbionts+Ocean*Salinity*Depth, data = plankonly, na.action = "na.fail")
summary(fullplankmodel)

fullplankmodel.dredge <- dredge(fullplankmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullplankmodel.dredge, delta < 4 | df == min(df)))

# Final 3/6/23
reducedplankmodel <- lm(D47ICDES ~ calcTiso*Dataset*carbsat + calcTiso*carbsat*Symbionts, data = plankonly, na.action = "na.fail")
summary(reducedplankmodel)
eta_squared(car::Anova(reducedplankmodel, type = 2))

check_collinearity(lm(D47ICDES ~ calcTiso+Dataset+carbsat+Symbionts, data = plankonly, na.action = "na.fail"))
# By dataset

ucla <- forams3[forams3$Dataset == "UCLA",]
uclaunscaled <- forams4[forams4$Dataset == "UCLA",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "ucla")

fulluclamodel <- lm(D47ICDES ~ calcTiso*carbsat*Salinity*Depth+BenthicPlanktic+Symbionts, data = ucla, na.action = "na.fail")
summary(fulluclamodel)

fulluclamodel.dredge <- dredge(fulluclamodel, cluster = clust, fixed = "calcTiso")
plot(subset(fulluclamodel.dredge, delta < 4 | df == min(df)))

# Final 3/6/23
reduceduclamodel <- lm(D47ICDES ~ calcTiso+carbsat*Depth+Salinity + Symbionts, data = ucla, na.action = "na.fail")
reduceduclamodel2 <- lm(D47ICDES ~ calcTiso+carbsat + Symbionts, data = ucla, na.action = "na.fail")
summary(reduceduclamodel)
anova(reduceduclamodel, reduceduclamodel2)
AIC(reduceduclamodel)
AIC(reduceduclamodel2) # 2 is worse
eta_squared(car::Anova(reduceduclamodel, type = 2))

reduceduclamodelunscaled <- lm(D47ICDES ~ calcTiso+carbsat*Depth+Salinity + Symbionts, data = uclaunscaled, na.action = "na.fail")
reduceduclamodelcoded <- lm(D47ICDES ~ calcTiso+carbsat*Depth+Salinity + Symbiontscoded, data = ucla, na.action = "na.fail")
summary(reduceduclamodelunscaled)

rescale.coefs <- function(beta,mu,sigma) {
  beta2 <- beta ## inherit names etc.
  beta2[-1] <- sigma[1]*beta[-1]/sigma[-1]
  beta2[1]  <- sigma[1]*beta[1]+mu[1]-sum(beta2[-1]*mu[-1])
  beta2
}
b1S <- coef(reduceduclamodel)
summary(reduceduclamodelunscaled)
summary(reduceduclamodel)$coefficients

sterr <- summary(reduceduclamodelcoded)$coefficients[,2]
mo <- forams4 %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(forams4)=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(forams4, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(forams4, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,0,1)

# If the model has an interaction:
cs. <- function(x) scale(x,center=TRUE,scale=TRUE)
m2 <- update(reduceduclamodelunscaled,. ~ . + carbsat:Depth)
m2.sc <- update(reduceduclamodel,. ~ . + I(cs.(carbsat*Depth)))
logLik(m2)-logLik(m2.sc)
#Calculate mean/sd of model matrix, dropping the first (intercept) value:
  
  X <- model.frame(m2)                                        
cm2 <- colMeans(X)[ -c(1,6)]
csd2 <- apply(X,2,sd)[-1]                                            
(cc2 <- rescale.coefs(fixef(m2.sc),mu=c(0,cm2),sigma=c(1,csd2)))
all.equal(unname(cc2),unname(fixef(m2)),tol=1e-3)  ## TRUE




#######

piasecki <- forams3[forams3$Dataset == "Piasecki (ETH 1, 3, 4)",]
piaseckiunscaled <- forams4[forams4$Dataset == "Piasecki (ETH 1, 3, 4)",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "piasecki")

fullpiaseckimodel <- lm(D47ICDES ~ calcTiso*carbsat*Salinity+Ocean+Depth, data = piasecki, na.action = "na.fail")
summary(fullpiaseckimodel)

fullpiaseckimodel.dredge <- dredge(fullpiaseckimodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullpiaseckimodel.dredge, delta < 5 | df == min(df)))

# No change 3/6/23
reducedpiaseckimodel <- lm(D47ICDES ~ calcTiso, data = piasecki, na.action = "na.fail")
summary(reducedpiaseckimodel)
eta_squared(car::Anova(reducedpiaseckimodel, type = 2)) # Eta squared instead of partial eta squared

b1S <- coef(reducedpiaseckimodel)
mo <- piaseckiunscaled %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(select_if(piaseckiunscaled, is.numeric))=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(piaseckiunscaled, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(piaseckiunscaled, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

##########
meinicke <- forams3[forams3$Dataset == "Meinicke 2019",]
meinickeunscaled <- forams4[forams4$Dataset == "Meinicke 2019",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "meinicke")

fullmeinickemodel <- lm(D47ICDES ~ calcTiso+carbsat+Salinity*Depth*Symbionts+Ocean, data = meinicke, na.action = "na.fail")
summary(fullmeinickemodel)

fullmeinickemodel.dredge <- dredge(fullmeinickemodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullmeinickemodel.dredge, delta < 4 | df == min(df)))

# Finalized 3/6/23 Confirmed that carbsat should not be in the model.
reducedmeinickemodel <- lm(D47ICDES ~ calcTiso + Ocean, data = meinicke, na.action = "na.fail")
summary(reducedmeinickemodel)
reducedmeinickemodel2 <- lm(D47ICDES ~ calcTiso, data = meinicke, na.action = "na.fail")
anova(reducedmeinickemodel, reducedmeinickemodel2) # No difference, default to simpler model
eta_squared(car::Anova(reducedmeinickemodel2, type = 2)) # Eta squared instead of partial eta squared

b1S <- coef(reducedmeinickemodel2)
mo <- meinickeunscaled %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(select_if(meinickeunscaled, is.numeric))=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(meinickeunscaled, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(meinickeunscaled, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

################
peral <- forams3[forams3$Dataset == "Peral 2018",]
peralunscaled <- forams4[forams4$Dataset == "Peral 2018",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "peral")

fullperalmodel <- lm(D47ICDES ~ calcTiso*carbsat*Depth*Symbionts*BenthicPlanktic, data = peral, na.action = "na.fail")
summary(fullperalmodel)

fullperalmodel.dredge <- dredge(fullperalmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullperalmodel.dredge, delta < 4 | df == min(df)))

# Confirmed best model 3/6/23
reducedperalmodel <- lm(D47ICDES ~ calcTiso + carbsat + Symbionts, data = peral, na.action = "na.fail")
summary(reducedperalmodel)
eta_squared(car::Anova(reducedperalmodel, type = 2)) # Eta squared instead of partial eta squared

b1S <- coef(reducedperalmodel)
mo <- peralunscaled %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(select_if(peralunscaled, is.numeric))=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(peralunscaled, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(peralunscaled, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

################

tripati <- forams3[forams3$Dataset == "Tripati",]
tripatiunscaled <- forams4[forams4$Dataset == "Tripati",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "tripati")

fulltripatimodel <- lm(D47ICDES ~ calcTiso*carbsat*Depth*Salinity*Symbionts, data = tripati, na.action = "na.fail")
summary(fulltripatimodel)

fulltripatimodel.dredge <- dredge(fulltripatimodel, cluster = clust, fixed = "calcTiso")
plot(subset(fulltripatimodel.dredge, delta < 4 | df == min(df)))
# No change 3/6/23
reducedtripatimodel <- lm(D47ICDES ~ calcTiso*Depth*Symbionts, data = tripati, na.action = "na.fail")
summary(reducedtripatimodel)
eta_squared(car::Anova(reducedtripatimodel, type = 2)) # Eta squared instead of partial eta squared

b1S <- coef(reducedtripatimodel)
mo <- tripatiunscaled %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(select_if(tripatiunscaled, is.numeric))=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(tripatiunscaled, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(tripatiunscaled, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

reducedtripatimodel

uclaall <- forams3[forams3$Dataset == "UCLA",]
clusterExport(clust, "uclaall")

fullUCLAmodel <- lm(D47ICDES ~ calcTiso+carbsat*Depth*Salinity*Symbionts+Ocean*BenthicPlanktic, data = uclaall, na.action = "na.fail")
summary(fullUCLAmodel)

fullUCLAmodel.dredge <- dredge(fullUCLAmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullUCLAmodel.dredge, delta < 4 | df == min(df)))

reducedUCLAmodel <- lm(D47ICDES ~ calcTiso+carbsat*Depth+Salinity+Symbionts, data = uclaall, na.action = "na.fail")
summary(reducedUCLAmodel)
reducedUCLAmodel2 <- lm(D47ICDES ~ calcTiso+carbsat, data = uclaall, na.action = "na.fail")
anova(reduceduclamodel, reducedUCLAmodel)

eta_squared(car::Anova(reducedUCLAmodel, type = 2))


# Benthics only
UCLAbenthic <- forams3[forams3$Dataset == "UCLA" & forams3$BenthicPlanktic == "Benthic",]
UCLAbenthicunscaled <- forams4[forams4$Dataset == "UCLA" & forams4$BenthicPlanktic == "Benthic",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "UCLAbenthic")

fullUCLAbenthicmodel <- lm(D47ICDES ~ calcTiso*carbsat+Depth*Salinity+Ocean+Region, data = UCLAbenthic, na.action = "na.fail")
summary(fullUCLAbenthicmodel)

fullUCLAbenthicmodel.dredge <- dredge(fullUCLAbenthicmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullUCLAbenthicmodel.dredge, delta < 0.4 | df == min(df)))

reducedUCLAbenthicmodel <- lm(D47ICDES ~ calcTiso+carbsat, data = UCLAbenthic, na.action = "na.fail")
reducedUCLAbenthicmodel2 <- lm(D47ICDES ~ calcTiso, data = UCLAbenthic, na.action = "na.fail")
summary(reducedUCLAbenthicmodel2) # No difference, go with simpler model
anova(reducedUCLAbenthicmodel, reducedUCLAbenthicmodel2)
eta_squared(car::Anova(reducedUCLAbenthicmodel, type = 2)) # Eta squared instead of partial eta squared
eta_squared(car::Anova(reducedUCLAbenthicmodel2, type = 2)) # Eta squared instead of partial eta squared

b1S <- coef(reducedUCLAbenthicmodel2)
mo <- UCLAbenthicunscaled %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(UCLAbenthicunscaled)=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(UCLAbenthicunscaled, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(UCLAbenthicunscaled, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

# Planktics only

UCLAplanktic <- forams3[forams3$Dataset == "UCLA" & forams3$BenthicPlanktic == "Planktic",]
UCLAplankticunscaled <- forams4[forams4$Dataset == "UCLA" & forams4$BenthicPlanktic == "Planktic",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "UCLAplanktic")

fullUCLAplankticmodel <- lm(D47ICDES ~ calcTiso*carbsat*Depth+Salinity*Symbionts+Ocean, data = UCLAplanktic, na.action = "na.fail")
summary(fullUCLAplankticmodel)

fullUCLAplankticmodel.dredge <- dredge(fullUCLAplankticmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullUCLAplankticmodel.dredge, delta < 2 | df == min(df)))
# No change 3/7/23
reducedUCLAplankticmodel <- lm(D47ICDES ~ calcTiso+carbsat+Symbionts, data = UCLAplanktic, na.action = "na.fail")
summary(reducedUCLAplankticmodel)
eta_squared(car::Anova(reducedUCLAplankticmodel, type = 2)) # Eta squared instead of partial eta squared

b1S <- coef(reducedUCLAplankticmodel)
mo <- UCLAplankticunscaled %>% 
        select(names(reducedUCLAplankticmodel[["model"]])) %>% 
          summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(UCLAplankticunscaled) =="D47ICDES")
UCLAplankticunscaled %>% select(names(reducedUCLAplankticmodel[["model"]])) %>% select_if(., is.numeric)


p.order <- 1:ncol(UCLAplankticunscaled %>% 
                    select(names(reducedUCLAplankticmodel[["model"]])) %>% 
                    select_if(., is.numeric)
                  )

m <- mo[p.order]
s <- apply(select_if(UCLAplankticunscaled[ , c(names(reducedUCLAplankticmodel[["model"]]))], is.numeric),2, sd, na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

# Starting point
b1S <- coef(reducedmeinickemodel2)
mo <- meinickeunscaled %>% summarize(across(where(is.numeric), ~mean(.x, na.rm = TRUE)))

icol <- which(colnames(select_if(meinickeunscaled, is.numeric))=="D47ICDES")
p.order <- c(icol,(1:ncol(select_if(meinickeunscaled, is.numeric)))[-icol])
m <- mo[p.order]
s <- apply(select_if(meinickeunscaled, is.numeric),2, sd,na.rm=TRUE)[p.order]
rescale.coefs(b1S,m,s)

# Benthic infaunal vs epifaunal ALL LABS


benthinfaunal <- forams3[forams3$Habitat == "Benthic, infaunal",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "benthinfaunal")


fullbenthinfaunalmodel <- lm(D47ICDES ~ calcTiso*Dataset*Depth*Salinity+carbsat+Ocean, data = benthinfaunal, na.action = "na.fail")
summary(fullbenthmodel)

fullbenthinfaunalmodel.dredge <- dredge(fullbenthinfaunalmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullbenthinfaunalmodel.dredge, delta < 5 | df == min(df)))

reducedbenthinfaunalmodel <- lm(D47ICDES ~ calcTiso+carbsat+Depth+Ocean, data = benthonly, na.action = "na.fail")
summary(reducedbenthinfaunalmodel)
reducedbenthinfaunalmodel2 <- lm(D47ICDES ~ calcTiso+carbsat, data = benthonly, na.action = "na.fail")
summary(reducedbenthinfaunalmodel2)
anova(reducedbenthinfaunalmodel, reducedbenthinfaunalmodel2)
AICc(reducedbenthinfaunalmodel)
AICc(reducedbenthinfaunalmodel2)

# They are the same, default to simpler model

eta_squared(car::Anova(reducedbenthinfaunalmodel2, type = 2)) # Eta squared instead of partial eta squared

# Benthic, epifaunal

benthepifaunal <- forams3[forams3$Habitat == "Benthic, epifaunal",]

clusterType <- if(length(find.package("snow", quiet = TRUE))) "SOCK" else "PSOCK"
clust <- try(makeCluster(getOption("cl.cores", 10), type = clusterType))

clusterExport(clust, "benthepifaunal")


fullbenthepifaunalmodel <- lm(D47ICDES ~ calcTiso*Dataset*Depth*Salinity+carbsat+Ocean, data = benthepifaunal, na.action = "na.fail")
summary(fullbenthepifaunalmodel)

fullbenthepifaunalmodel.dredge <- dredge(fullbenthepifaunalmodel, cluster = clust, fixed = "calcTiso")
plot(subset(fullbenthepifaunalmodel.dredge, delta < 5 | df == min(df)))

reducedbenthepifaunalmodel <- lm(D47ICDES ~ calcTiso + Depth, data = benthepifaunal, na.action = "na.fail")
summary(reducedbenthepifaunalmodel)

eta_squared(car::Anova(reducedbenthepifaunalmodel, type = 2)) # Eta squared instead of partial eta squared

##############
# Output model tables using flextable
set_flextable_defaults(
  font.size = 10, theme_fun = theme_vanilla, font.family = "arial",
  padding = 2,
  background.color = "#FFFFFF",
  split = FALSE)


data.frame(summary(reducedbenthepifaunalmodel))
as_flextable(reducedbenthepifaunalmodel)
library(officer)
sect_properties <- prop_section(
  page_size = page_size(
    orient = "landscape",
    width = 8.5, height = 11
  ),
  type = "continuous",
  page_margins = page_mar()
)


save_as_docx(
  `UCLA full dataset` = colformat_double(as_flextable(eta_squared(car::Anova(reduceduclamodel, type = 2))[,c(1,2,4)]), digits = 2),
  `UCLA benthic foraminifera` = colformat_double(as_flextable(eta_squared(car::Anova(reducedUCLAplankticmodel, type = 2))[,c(1,2,4)]), digits = 2),
  `UCLA  planktic foraminifera` = colformat_double(as_flextable(eta_squared(car::Anova(reducedUCLAbenthicmodel2, type = 2))[,c(1,2,4)]), digits = 2),
  `Piasecki full dataset` = colformat_double(as_flextable(eta_squared(car::Anova(reducedpiaseckimodel, type = 2))[,c(1,2,4)]), digits = 2),
  `Meinicke full dataset` = colformat_double(as_flextable(eta_squared(car::Anova(reducedmeinickemodel2, type = 2))[,c(1,2,4)]), digits = 2),
  `Peral full dataset` = colformat_double(as_flextable(eta_squared(car::Anova(reducedperalmodel, type = 2))[,c(1,2,4)]), digits = 2),
  `Tripati full dataset` = colformat_double(as_flextable(eta_squared(car::Anova(reducedtripatimodel, type = 2))), digits = 2),
  `All planktic foraminifera` = colformat_double(as_flextable(eta_squared(car::Anova(reducedplankmodel, type = 2))[,c(1,2,4)]), digits = 2),
  `All benthic foraminifera` = colformat_double(as_flextable(eta_squared(car::Anova(reducedbenthmodel, type = 2))[,c(1,2,4)]), digits = 2),
  `Full foraminifera dataset` = colformat_double(as_flextable(eta_squared(car::Anova(reducedmodel, type = 2))[,c(1,2,4)]), digits = 2),
  path = "SI effect size tables.docx", 
  pr_section = prop_section(
    page_size = page_size(
      orient = "portrait",
      width = 8.5, height = 11
    ),
    type = "continuous",
    page_margins = page_mar()
  ),
  align = "left"
)

save_as_docx(
  `UCLA full dataset` = as_flextable(reduceduclamodel), 
  `UCLA benthic foraminifera` = as_flextable(reducedUCLAplankticmodel),
  `UCLA  planktic foraminifera` = as_flextable(reducedUCLAbenthicmodel2),
  `Piasecki full dataset` = as_flextable(reducedpiaseckimodel),
  `Meinicke full dataset` = as_flextable(reducedmeinickemodel2),
  `Peral full dataset` = as_flextable(reducedperalmodel),
  `Tripati full dataset` = as_flextable(reducedtripatimodel),
  `All planktic foraminifera` = as_flextable(reducedplankmodel),
  `All benthic foraminifera` = as_flextable(reducedbenthmodel),
  `Full foraminifera dataset` = as_flextable(reducedmodel),
  path = "SI model tables.docx", 
  pr_section = sect_properties,
  align = "left"
)

#######################
# Test of the smatr package for deming regression model comparison

library(deming)
library(smatr)
# Pull from forams data
uclaplank <- forams[forams$Dataset == "UCLA" & forams$BenthicPlanktic == "Planktic",]
uclaplankmod <- deming(D47ICDES ~ calcTiso, data = uclaplank, xstd = calcTiso.SE, ystd = D47ICDES.SE, na.action = na.omit)
uclaplankmod

uclabenth <- forams[forams$Dataset == "UCLA" & forams$BenthicPlanktic == "Benthic",]
uclabenthmod <- deming(D47ICDES ~ calcTiso, data = uclabenth, xstd = calcTiso.SE, ystd = D47ICDES.SE, na.action = na.omit)
uclabenthmod

smatr::slope.com(uclabenthmod, uclabenthmod, BenthicPlanktic)
with(leaflife, slope.com(log10(longev), log10(lma), site,  method='MA', alpha=0.05))
with(leaflife, slope.com(log10(longev), log10(lma), site))
anova(model.matrix(uclabenthmod), uclaplankmod)
