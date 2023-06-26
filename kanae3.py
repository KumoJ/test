import cv2

#元画像の読み込み
img3 = cv2.imread('data/kanae.png')

img2 = cv2.cvtColor(img3,cv2.COLOR_BGR2GRAY)

img = cv2.resize(img2, dsize=(54,96))

#画像の縦の画素数、横の画素数を取得
h,w = img.shape


#３値化に使う低い側の閾値を決める
th_low_value=int(input('1つ目の閾値を入力してください。:'))
#３値化に使う高い側の閾値を決める
th_high_value=int(input('2つ目の数字を入力してください。:'))
#3値化の白黒以外の色
gray=100

#画像の３値化
for i in range(h):
    for j in range(w):
        if img[i,j]<th_low_value:
            img[i,j]=0
        elif th_low_value <= img[i,j] <=th_high_value:
            img[i,j]=gray
        else :
            img[i,j]=255

#3値化した画像の表示
cv2.imshow("Kanae3",img)
cv2.waitKey(0)
cv2.destroyAllWindows()

cv2.imwrite('data/kanae3.png',img)