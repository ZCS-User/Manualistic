# library
from PIL import Image


def unifilare(serie, fase, sonda):# opening up of images
    img_prodotto = Image.open("./img/prodotto/"+serie+".png")
    # img_lineaO = Image.open("./img/misc/lineaO.png")
    # img_lineaV = Image.open("./img/misc/lineaV.png")
    # img_rete = Image.open("./img/misc/SchemaRete.png")
    if sonda == 'METER':
        if '1' in fase:
            img_sonda = Image.open("./img/misc/ZSM-METER-DDSU.png")
        else:
            img_sonda = Image.open("./img/misc/ZSM-METER-DTSU.png")
    elif sonda in ['TA', 'CT']:
        img_sonda = Image.open("./img/misc/CurrentSensor.png")
    elif sonda in 'ENERCLICK':
        img_sonda = Image.open("./img/misc/COMBOX.png")
    img_schema = Image.open("./img/misc/SchemaTot.png")
    #img1 = Image.open("./img/zcs.png")
    img_prodotto = img_prodotto.resize((130, 120))
    a = img_schema.size[0]  # +img_rete.size[0]+img_lineaO.size[0]
    b = img_schema.size[1]  # max(img_prodotto.size[1], img_rete.size[1])


    img = Image.new("RGB", (a, b), "white")
    img.paste(img_schema, (0, 0))
    img.paste(img_prodotto, (0, 0))
    if sonda == 'METER':
        img.paste(img_sonda, (int(0.53*img_schema.size[0]), int(img_schema.size[1]-img_sonda.size[1])))
    elif sonda in ['TA', 'CT']:
        img_sonda = img_sonda.resize((60, 105))
        img.paste(img_sonda, (int(0.55*img_schema.size[0]), int(0.1*img_schema.size[1]/3)))
    elif sonda in 'ENERCLICK':
        img_sonda = img_sonda.resize((150, 100))
        img.paste(img_sonda, (int(0.45*img_schema.size[0]), int(img_schema.size[1]-img_sonda.size[1])))
    img.save('./0inj/img/Schema_'+serie+'_'+sonda+'.png')
