fname = "C:\\abc\\input.dat"
con = file(fname, "rb")
colname = readBin(con, raw(),  file.info(fname)$size)
close(con)

nn=972000
x <- colname[1:nn]
outBytes =numeric(nn * 6)

for (i in 1:nn) {
  xx = as.integer(x[i])
  j = i - 1
  outBytes[j * 6 + 1] = outBytes[j * 6 + 4] = xx
  outBytes[j * 6 + 2] = outBytes[j * 6 + 5] = (xx * 3) %% 256
  outBytes[j * 6 + 3] = outBytes[j * 6 + 6] = (xx * 5) %% 256
  
}

y = file("C:\\outdata_R.bin","wb")
writeBin(as.raw(outBytes), y)
close(y)
