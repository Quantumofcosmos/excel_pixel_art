import pygame, sys ,xlsxwriter, errno
from os import remove
from PIL import Image
from skimage import io


basewidth = 900

try:
    argu=sys.argv[1]
    img = Image.open(sys.argv[1])
except:
    print("Working with default image")
    argu="Taj"
    img = Image.open('taj.jpeg')
wpercent = (basewidth/float(img.size[0]))
hsize = int((float(img.size[1])*float(wpercent)))
img = img.resize((basewidth,hsize), Image.ANTIALIAS)
img.save('topyforcrop.jpg')
pygame.init()
pygame.display.set_caption('Click anywhere on the image to start cropping..')

def displayImage(screen, px, topleft, prior):
    # ensure that the rect always has positive width, height
    x, y = topleft
    width =  pygame.mouse.get_pos()[0] - topleft[0]
    height = pygame.mouse.get_pos()[1] - topleft[1]
    if width < 0:
        x += width
        width = abs(width)
    if height < 0:
        y += height
        height = abs(height)

    # eliminate redundant drawing cycles (when mouse isn't moving)
    current = x, y, width, height
    if not (width and height):
        return current
    if current == prior:
        return current

    # draw transparent box and blit it onto canvas
    screen.blit(px, px.get_rect())
    im = pygame.Surface((width, height))
    im.fill((128, 128, 128))
    pygame.draw.rect(im, (32, 32, 32), im.get_rect(), 1)
    im.set_alpha(128)
    screen.blit(im, (x, y))
    pygame.display.flip()

    # return current box extents
    return (x, y, width, height)

def setup(path):
    px = pygame.image.load(path)
    screen = pygame.display.set_mode( px.get_rect()[2:] )
    screen.blit(px, px.get_rect())
    pygame.display.flip()
    return screen, px

def mainLoop(screen, px):
    topleft = bottomright = prior = None
    n=0
    while n!=1:
        for event in pygame.event.get():
            if event.type == pygame.MOUSEBUTTONUP:
                if not topleft:
                    topleft = event.pos
                else:
                    bottomright = event.pos
                    n=1
        if topleft:
            prior = displayImage(screen, px, topleft, prior)
    return ( topleft + bottomright )

def exceler():
#    baswidth = 900
#    img = Image.open(master)
#    wpercent = (baswidth/float(img.size[0]))
#    hsize = int((float(img.size[1])*float(wpercent)))
#    img = img.resize((baswidth,hsize), Image.ANTIALIAS)
#    img.save('cropin.jpg')
    img = io.imread(output_loc)
    print(img.shape)
    sheetname=argu+".xlsx"
    workbook = xlsxwriter.Workbook(sheetname)
    worksheet = workbook.add_worksheet()
	#worksheet.set_column('A:DW', 1.5)
	#worksheet.set_row(1:129, 10)

    for row in list(range(0,img.shape[0])):
    #    worksheet.set_row(row, 1.5)
        for col in list(range(0,img.shape[1])):
            r,g,b=img[row,col].tolist()
            hexx='#'+hex(r).replace('x','0')[-2:]+hex(g).replace('x','0')[-2:]+hex(b).replace('x','0')[-2:]
            cell_format = workbook.add_format()
            cell_format.set_bg_color(hexx)
            worksheet.write(row,col, ' ',cell_format)
	#       break
	#    break 21 15 15 18 14 13

    for row in list(range(0,img.shape[0])):
        worksheet.set_row(row, 25)

    colszrange='A:'+xlsxwriter.utility.xl_col_to_name(img.shape[1])
    worksheet.set_column(colszrange, 4)

    workbook.close()

def silentremove(filename):
    try:
        remove(filename)
    except OSError as e: # this would be "except OSError, e:" before Python 2.6
        if e.errno != errno.ENOENT: # errno.ENOENT = no such file or directory
            raise

if __name__ == "__main__":
    input_loc = 'topyforcrop.jpg'
    output_loc = 'cropout.jpg'
    screen, px = setup(input_loc)
    left, upper, right, lower = mainLoop(screen, px)

    # ensure output rect always has positive width, height
    if right < left:
        left, right = right, left
    if lower < upper:
        lower, upper = upper, lower
    im = Image.open(input_loc)
    im = im.crop(( left, upper, right, lower))
    pygame.display.quit()
    im.save(output_loc)
    print("Done processing loading into excel file")
    exceler()
    print("cleaning up behind process")
    silentremove('cropout.jpg')
    silentremove('topyforcrop.jpg')
    pygame.quit()
    print("Done")
