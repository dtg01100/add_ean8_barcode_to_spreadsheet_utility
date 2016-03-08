__version__ = "$Revision: 0.4$"

codeLength = { "EAN13": 13,
               "EAN8": 8,
               "UPCA": 12,
               "UPCE": 8 }

def __lenCheck(chk, _type='EAN13'):
    return (len(chk) == codeLength[_type]) or \
       (len(chk) == codeLength[_type]-1)


def __sumDigits(chk, start=0, end=1, step=2, mult=1):
    return reduce(lambda x, y: int(x)+int(y), list(chk[start:end:step])) * mult

def eanCheckDigit(chk, code='EAN13'):
    """Returns the checksum digit of an EAN-13/8 code"""
    if chk.isdigit() and __lenCheck(chk):
        if code == 'EAN13':
            m0=1
            m1=3
        elif code == 'EAN8':
            m0=3
            m1=1
        else:
            return None

        _len = codeLength[code]-1
        t = 10 - (( __sumDigits(chk, start=0, end=_len, mult=m0) + \
                    __sumDigits(chk, start=1, end=_len, mult=m1) ) %10 ) %10

        if t == 10:
            return 0
        else:
            return t

    return None

def ean13Valid(chk):
    """Verify if code is valid EAN13 barcode.  Returns True|False"""
    return chk.isdigit() and __lenCheck(chk) and \
           (int(chk[-1]) == eanCheckDigit(chk))

def ean8CheckDigit(chk):
    """Returns the checksum digit of an EAN8 code"""
    return eanCheckDigit(chk, code='EAN8')

def ean8Valid(chk):
    """Verify if code is valid EAN8 barcode. Returns True|False"""
    if chk.isdigit() and len(chk) == codeLength["EAN8"]:
        return int(chk[-1]) == ean8CheckDigit(chk)
    return False

# UPC codes below

def upcaCheckDigit(chk):
    if chk is not None:
        return eanCheckDigit('0'+chk)
    return None

def upcaValid(chk):
    if chk is not None:
        return ean13Valid('0'+chk)
    return False

def upca2e(chk):
    t = None
    if chk.isdigit() and __lenCheck(chk, 'UPCA'):
        if '012'.find(chk[3]) >= 0 and chk[4:8] == '0000':
            t=chk[:3]+chk[8:11]+chk[3]
        elif chk[4:9] == '00000':
            t=chk[:4]+chk[9:11]+'3'
        elif chk[5:10] == '00000':
            t = chk[:5]+chk[10]+'4'
        elif '5678'.find(chk[10]) >= 0 and chk[6:10] == '0000':
            t=chk[:6]+chk[10]
        else:
            t=None

        # Check digit
        if t is not None:
            if upcaValid(chk):
                t=t+chk[-1]
            elif len(chk) == codeLength["UPCA"]-1:
                t=t+str(upcaCheckDigit(chk))
            else:
                t=None
    return t

def upce2a(chk):
    t=None
    if chk.isdigit() and __lenCheck(chk, 'UPCE'):
        if '012'.find(chk[6]) >= 0:
            t=chk[:3]+chk[6]+'0000'+chk[3:6]
        elif chk[6] == '3':
            t=chk[:4]+'00000'+chk[4:6]
        elif chk[6] == '4':
            t=chk[:5]+'00000'+chk[5]
        elif '5678'.find(chk[6]) >= 0:
            t=chk[:6]+'0000'+chk[6]
        else:
            t=None

        if t is not None:
            if len(chk) == codeLength["UPCE"] - 1:
                t=t+str(upcaCheckDigit(t))
            elif len(chk) == codeLength['UPCE'] and \
                 int(chk[-1]) == upcaCheckDigit(t):
                t=t+chk[-1]
            else:
                t=None
    return t

def upceValid(chk):
    return len(chk) == codeLength["UPCE"] and upcaValid(upce2a(chk))

def upceCheckDigit(chk):
    if chk is not None:
        return upcaCheckDigit(upce2a(chk))
    return None
