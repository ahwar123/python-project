from skimage.filters import threshold_mean
from skimage import measure
from skimage import transform as trans
import matplotlib.pyplot as plt
import numpy as np
import math
import xlsxwriter
import skimage as ski
import skimage.io as skio
from skimage import filters
import matplotlib.patches as patches
import glob


i=1
j=1
q=1
counter=1
book = xlsxwriter.Workbook('rice_grain_detail_sheet.xlsx')
worksheet = book.add_worksheet()
bold = book.add_format({'bold': True})

worksheet.write(0, 1, 'Image name',bold)
worksheet.write(0, 2, 'Length',bold)
worksheet.write(0, 3, 'Width',bold)
worksheet.write(0, 4, 'Mean',bold)
worksheet.write(0, 5, 'Min',bold)
worksheet.write(0, 6, 'Max',bold)
worksheet.write(0, 7, 'Quality',bold)

path = 'images/*.jpg'
files=glob.glob(path)

for ImageFileName in files:
    print(ImageFileName)

    img = skio.imread(ImageFileName, as_grey=True)


    t= threshold_mean(img)

    thresh_img= img > t

    label_img = measure.label(thresh_img)

    props = measure.regionprops(label_img)

    #
    #
    # plt.figure(1)
    # plt.subplot(131)
    # plt.imshow(img,cmap='gray')
    # plt.axis('off')
    # plt.subplot(132)
    # plt.imshow(label_img, cmap='spectral')
    # plt.axis('off')
    # b=plt.subplot(133)
    # plt.imshow(label_img, cmap='gray')
    # plt.axis('off')


    for prop in props:
        if(prop.area>9000 and prop.eccentricity>0.3):


            # minr, minc, maxr, maxc = prop['BoundingBox']
            minr, minc, maxr, maxc = prop.bbox
            croped_img=label_img[minr:maxr,minc:maxc]


            bx = (minc, maxc, maxc, minc, minc)
            by = (minr, minr, maxr, maxr, minr)


            ax= plt.plot(bx, by, '-y', linewidth=1)
            cprops= measure.regionprops(croped_img)
            # plt.figure(2)
            # # plt.subplot(i)
            #
            # plt.subplot(111)
            # plt.imshow(croped_img, cmap='gray')

            for cprop in cprops:
                if cprop.area > 1000 and prop.eccentricity > 0.9:

                    x0 = cprop['Centroid'][1]
                    y0 = cprop['Centroid'][0]
                    x1 = x0 + math.cos(cprop['Orientation']) * 0.5 * cprop['MajorAxisLength']
                    y1 = y0 - math.sin(cprop['Orientation']) * 0.5 * cprop['MajorAxisLength']
                    x2 = x0 - math.sin(cprop['Orientation']) * 0.5 * cprop['MinorAxisLength']
                    y2 = y0 - math.cos(cprop['Orientation']) * 0.5 * cprop['MinorAxisLength']


                    plt.plot((x0, x1), (y0, y1), '-y', linewidth=1)     #length radius
                    plt.plot((x0, x2), (y0, y2), '-y', linewidth=1)     #width radius
                    plt.plot(x0, y0, '-g', markersize=15)
                    angle=cprop.orientation

                    croped_img = label_img[minr:maxr, minc:maxc]

                    angleRad = math.atan2(y0- y1, x0 - x1)         #get angle between x and y axis
                    a = math.degrees(angleRad)


                    final_img = trans.rotate(croped_img, a+90, resize=True)


                    #get length of the grain
                    p1 = [x0,x1]
                    p2 = [y0, y1]
                    length = math.sqrt( ( (p2[0] - p1[0]) **2) + ( (p2[1] - p1[1]) **2) )
                    length=length*2

                    #get width of the grain
                    w1 = [x0, x2]
                    w2 = [y0, y2]
                    width = math.sqrt( ( (w2[0] - w1[0]) **2) + ( (w2[1] - w1[1]) **2) )
                    width = width * 2

                    arr = np.asarray(np.double(final_img))  # convert image as array

                    min = arr.min()       #minumum
                    max = arr.max()       #maximum
                    mean= arr.mean()      #mean

                    print('Length=' + str(length))
                    print('Width=' + str(width))
                    print('mean='+str(mean))
                    print('max='+str(max))
                    print('min='+str(min))

                    plt.imsave('rice_grain_images\\'+ str(i) + '.tif',final_img)    #save image in folder

                    # Write grain info in the excel file
                    worksheet.write(i,j, str(i)+'.tif')
                    worksheet.write(i,j+1, length)
                    worksheet.write(i,j+2, width)
                    worksheet.write(i,j+3, mean)
                    worksheet.write(i,j+4, min)
                    worksheet.write(i,j+5, max)

                    #check for the quality
                    if q==1:
                        worksheet.write(i,j+6, 'C9')
                    elif q==2:
                        worksheet.write(i,j+6, 'Super Basmati')
                    elif q==3:
                        worksheet.write(i,j+6, 'Supri')
                    elif q==4:
                        worksheet.write(i,j+6, '386')
                    elif q==5:
                        worksheet.write(i,j+6, '1121 Steamed')
                    i+=1

    if counter%5==0:
        q+=1
    counter += 1
                    # plt.figure(3)
                    # # plt.subplot(i)
                    #
                    # plt.subplot(111)
                    # plt.imshow(final_img, cmap='gray')


                    # a = b.add_patch




                    # plt.gray()
                    # plt.axis((0, 4000, 3000, 0))
                    #
                    #
                    #
                    # plt.show()

book.close()
