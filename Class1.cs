

# カルマンフィルタの考え方
# Logics of Blue

# --------------------------------------------------------------------------------------
# ライブラリを使わないカルマンフィルタの実装
# --------------------------------------------------------------------------------------


localLevelModel <- function(y, xPre, pPre, sigmaW, sigmaV) {
  # お手製のカルマンフィルタ関数
  
  # 状態の予測(ローカルレベルモデルなので、予測値は、前期の値と同じ)
  xForecast <- xPre
  
  # 状態の予測誤差の分散
  pForecast <- pPre + sigmaW
  
  # カルマンゲイン
  kGain <- pForecast / (pForecast + sigmaV)
  
  # カルマンゲインを使って補正された状態
  xFiltered <- xForecast + kGain * (y - xForecast)
  
  # 補正された状態の予測誤差の分散
  pFiltered <- (1 - kGain) * pForecast
  
  # 結果の格納
  result <- data.frame(
    xFiltered = xFiltered, 
    pFiltered = pFiltered
  )
  
  return(result)
}


# ナイル川流量データで計算
Nile

# サンプルサイズ
N <- length(Nile)

# 状態の推定値
x <- numeric(N)	

# 状態の予測誤差の分散
P <- numeric(N)	


# 「状態」の初期値は0とします
x <- c(0, x)
x

# 「状態の予測誤差の分散」の初期値は1000にします
P <- c(1000, P)
P




# カルマンフィルタの逐次計算を行う
for(i in 1:N) {
  kekka <- localLevelModel(Nile[i], x[i], P[i], sigmaW = 1000, sigmaV = 10000)
  x[i + 1] <- kekka$xFiltered
  P[i + 1] <- kekka$pFiltered
}

# 結果の図示
year <- 1871:1970
plot(
  as.numeric(Nile) ~ year,
  main="お手製カルマンフィルタ関数の結果",
  ylab="ナイル川の流量",
  type="l", lty=2
)

lines(x[-1] ~ year, type="l", col=2)

legend(
  "topright",
  legend = c("観測値", "フィルタリングされた結果"),
  lty = c(2,1), col = c(1,2)
)



# --------------------------------------------------------------------------------------
# パラメタを変えたときの挙動
# --------------------------------------------------------------------------------------

calcState <- function(data, x0, P0, sigmaW, sigmaV){
  # カルマンフィルタを使って、状態を一気に推定する
  
  # サンプルサイズ
  N <- length(data)
  
  # 状態の予測誤差の分散
  P <- numeric(N)	
  
  # 「状態の予測誤差の分散」の初期値の設定
  P <- c(P0, P)
  
  # 状態の推定値
  x <- numeric(N)	
  
  # 「状態」の初期値の設定
  x <- c(x0, x)
  
  # カルマンフィルタの逐次計算を行う
  for(i in 1:N) {
    kekka <- localLevelModel(Nile[i],x[i],P[i], sigmaW = sigmaW, sigmaV = sigmaV)
    x[i + 1] <- kekka$xFiltered
    P[i + 1] <- kekka$pFiltered
  }
  
  # 推定された状態を返す
  return(x[-1])
  
}


# 最初に計算したパラメタと同じ
x1 <- calcState(data=Nile, x0=0,    P0=1000,   sigmaW=1000, sigmaV=10000)

# 「状態の予測誤差の分散」の初期値を増やした
x2 <- calcState(data=Nile, x0=0,    P0=100000, sigmaW=1000, sigmaV=10000)

# 状態方程式のノイズの分散をとても小さくした
x3 <- calcState(data=Nile, x0=1000, P0=0.1,    sigmaW=0.001,sigmaV=1000000)

# 観測方程式におけるノイズの分散をとても小さくした
x4 <- calcState(data=Nile, x0=1000, P0=100000, sigmaW=10000, sigmaV=100)

# 結果
year <- 1871:1970
plot(
  as.numeric(Nile) ~ year,
  main="パラメタを変えた結果",
  ylab="ナイル川の流量",
  type="l", lty=2, lwd=2
)

lines(x1 ~ year, type="l", col=2)
lines(x2 ~ year, type="l", col=6)
lines(x3 ~ year, type="l", col=4)
lines(x4 ~ year, type="l", col=5)

legend(
  "topright",
  legend = c("観測値", "x1:P0小", "x2:P0大", "x3:観測誤差大", "x4:観測誤差小"),
  lty = c(2,1,1,1,1), col = c(1,2,6,4,5)
)


