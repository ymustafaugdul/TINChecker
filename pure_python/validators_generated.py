from __future__ import annotations

import re
from datetime import date


class Ref:
    def __init__(self, value=None):
        self.value = value


def ensure_ref(value):
    return value if isinstance(value, Ref) else Ref(value)


def unwrap(value):
    return value.value if isinstance(value, Ref) else value


class VbaArray:
    def __init__(self, lower, upper, default=None):
        self.lower = int(lower)
        self.upper = int(upper)
        self.data = {index: default for index in range(self.lower, self.upper + 1)}

    def __getitem__(self, index):
        return self.data[int(unwrap(index))]

    def __setitem__(self, index, value):
        self.data[int(unwrap(index))] = unwrap(value)

    def __call__(self, index):
        return self.__getitem__(index)


class VbaLiteralArray:
    def __init__(self, values):
        self.data = tuple(unwrap(value) for value in values)
        self.lower = 0
        self.upper = len(self.data) - 1

    def __getitem__(self, index):
        return self.data[int(unwrap(index))]

    def __call__(self, index):
        return self.__getitem__(index)


class VbaDictionary(dict):
    def Add(self, key, value):
        self[unwrap(key)] = unwrap(value)

    def __call__(self, key):
        return self[unwrap(key)]


class VbaRegExp:
    def __init__(self):
        self.Pattern = ''
        self.IgnoreCase = False
        self.Global = False

    def Test(self, value):
        flags = re.IGNORECASE if self.IgnoreCase else 0
        return re.fullmatch(self.Pattern, str(unwrap(value)), flags) is not None


class _InvalidDate:
    year = -1
    month = -1
    day = -1


class _VbaError:
    pass


class _Application:
    @staticmethod
    def Match(value, values, match_type=0):
        values = unwrap(values)
        value = unwrap(value)
        try:
            if isinstance(values, VbaLiteralArray):
                return values.data.index(value) + 1
            if isinstance(values, VbaArray):
                for index in range(values.lower, values.upper + 1):
                    if values(index) == value:
                        return index
            return list(values).index(value) + 1
        except (ValueError, TypeError):
            return _VbaError()


Application = _Application()


def CreateObject(name):
    if name == 'Scripting.Dictionary':
        return VbaDictionary()
    if name == 'VBScript.RegExp':
        return VbaRegExp()
    raise ValueError(f'Unsupported CreateObject: {name}')


def Len(value):
    return len(str(unwrap(value)))


def Mid(value, start, length=None):
    value = str(unwrap(value))
    start = int(unwrap(start)) - 1
    return value[start:] if length is None else value[start:start + int(unwrap(length))]


def Left(value, length):
    return str(unwrap(value))[: int(unwrap(length))]


def Right(value, length):
    return str(unwrap(value))[-int(unwrap(length)) :]


def Replace(value, find, repl):
    return str(unwrap(value)).replace(str(unwrap(find)), str(unwrap(repl)))


def Trim(value):
    return str(unwrap(value)).strip()


def UCase(value):
    return str(unwrap(value)).upper()


def LCase(value):
    return str(unwrap(value)).lower()


def IsNumeric(value):
    value = str(unwrap(value)).strip()
    if value == '':
        return False
    try:
        float(value)
        return True
    except ValueError:
        return False


def CInt(value):
    return int(float(unwrap(value)))


def CLng(value):
    return int(float(unwrap(value)))


def StringFunc(length, character):
    return str(unwrap(character)) * int(unwrap(length))


def InStr(*args):
    if len(args) == 2:
        start, string1, string2 = 1, args[0], args[1]
    else:
        start, string1, string2 = args
    idx = str(unwrap(string1)).find(str(unwrap(string2)), int(unwrap(start)) - 1)
    return idx + 1 if idx >= 0 else 0


def Asc(value):
    return ord(str(unwrap(value))[0])


def Chr(value):
    return chr(int(unwrap(value)))


def Val(value):
    value = str(unwrap(value)).strip()
    match = re.match(r'^[+-]?(?:\d+(?:\.\d*)?|\.\d+)', value)
    if not match:
        return 0
    number = match.group(0)
    return float(number) if '.' in number else int(number)


def CStr(value):
    return str(unwrap(value))


def Array(*values):
    return VbaLiteralArray(values)


vbCrLf = "\r\n"


def UBound(value, dimension=None):
    value = unwrap(value)
    if isinstance(value, (VbaArray, VbaLiteralArray)):
        return value.upper
    if isinstance(value, (list, tuple)):
        return len(value) - 1
    raise ValueError('Unsupported UBound target')


def IsError(value):
    return isinstance(unwrap(value), _VbaError)


def IsDate(value):
    value = unwrap(value)
    return isinstance(value, date) and getattr(value, "year", -1) > 0


def IsAllDigits(value):
    return str(unwrap(value)).isdigit()


def CalculateCheckDigit(value, weights):
    clean_value = str(unwrap(value))
    total = 0
    for index in range(0, UBound(weights) + 1):
        total += CInt(Mid(clean_value, index + 1, 1)) * weights(index)
    return total % 11


def Format(value, fmt):
    fmt = str(unwrap(fmt))
    value = unwrap(value)
    if fmt == '00':
        return f'{int(value):02d}'
    raise ValueError(f'Unsupported Format pattern: {fmt}')


def DateSerial(year_value, month_value, day_value):
    try:
        return date(int(unwrap(year_value)), int(unwrap(month_value)), int(unwrap(day_value)))
    except Exception:
        return _InvalidDate()


def day(value):
    value = unwrap(value)
    return getattr(value, 'day', -1)


def month(value):
    value = unwrap(value)
    return getattr(value, 'month', -1)


def year(value):
    value = unwrap(value)
    return getattr(value, 'year', -1)


def vba_range(start, end, step=1):
    start = int(unwrap(start))
    end = int(unwrap(end))
    step = int(unwrap(step))
    return range(start, end + 1, step) if step > 0 else range(start, end - 1, step)


def vba_like(value, pattern):
    value = str(unwrap(value))
    pattern = str(unwrap(pattern))
    regex = ''
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == '#':
            regex += r'\d'
        elif ch == '?':
            regex += '.'
        elif ch == '*':
            regex += '.*'
        elif ch == '[':
            end = pattern.find(']', i)
            if end == -1:
                regex += re.escape(ch)
            else:
                regex += pattern[i:end + 1]
                i = end
        else:
            regex += re.escape(ch)
        i += 1
    return re.fullmatch(regex, value) is not None


class VBA:
    Date = date.today()
    day = staticmethod(day)
    month = staticmethod(month)
    year = staticmethod(year)


def validate_entry(country_code, vkn):
    formatInfo = Ref('')
    errorMsg = Ref('')
    is_valid = ValidateCountryDispatch(vkn, country_code, formatInfo, errorMsg)
    return {
        'countryCode': UCase(Trim(country_code)),
        'vkn': Trim(vkn),
        'isValid': bool(is_valid),
        'formatInfo': formatInfo.value,
        'errorMsg': errorMsg.value,
    }


def IsAllNumeric(str):
    _result = False
    i = 0
    if Len(str) == 0:
        _result = False
        return _result
    for i in vba_range(1, Len(str), 1):
        if not (vba_like(Mid(str, i, 1), "[0-9]")):
            _result = False
            return _result
    _result = True
    return _result

def IsAlphanumeric(str):
    _result = False
    i = 0
    _result = True
    for i in vba_range(1, Len(str), 1):
        if not vba_like(Mid(str, i, 1), "[A-Za-z0-9]"):
            _result = False
            return _result
    return _result

def IsAllLetters(str):
    _result = False
    i = 0
    _result = True
    for i in vba_range(1, Len(str), 1):
        if not vba_like(Mid(str, i, 1), "[A-Za-z]"):
            _result = False
            return _result
    return _result

def IsLetter(char):
    _result = False
    if Len(char) != 1:
        _result = False
    else:
        _result = vba_like(char, "[A-Z]")
    return _result

def IsValidDay(day, month, year):
    _result = False
    maxDay = 0
    if month == 1 or month == 3 or month == 5 or month == 7 or month == 8 or month == 10 or month == 12:
        maxDay = 31
    elif month == 4 or month == 6 or month == 9 or month == 11:
        maxDay = 30
    elif month == 2:
        if IsLeapYear(year):
            maxDay = 29
        else:
            maxDay = 28
    else:
        _result = False
        return _result
    if day >= 1 and day <= maxDay:
        _result = True
    else:
        _result = False
    return _result

def IsLeapYear(year):
    _result = False
    if (year % 4 == 0):
        if (year % 100 == 0):
            if (year % 400 == 0):
                _result = True
            else:
                _result = False
        else:
            _result = True
    else:
        _result = False
    return _result

def ValidateAfghanistanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Afganistan TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10:
        errorMsg.value = "Afganistan TIN'i 10 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateAlbaniaVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(vkn) != 10:
        errorMsg.value = "Arnavutluk VKN'si 10 karakter olmalıdır."
        return _result
    _result = True
    return _result

def ValidateAlgeriaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Cezayir TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 15 and Len(cleanTIN) != 20:
        errorMsg.value = "Cezayir TIN'i 15 veya 20 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateAndorraVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    firstChar = ""
    numericPart = ""
    controlLetter = ""
    validFirstChars = ""
    numValue = 0
    validRange = False
    vkn = UCase(Replace(Replace(Trim(vkn), "-", ""), " ", ""))
    if Len(vkn) != 8:
        errorMsg.value = "Andorra VKN'si 8 karakter olmalıdır."
        return _result
    firstChar = Left(vkn, 1)
    numericPart = Mid(vkn, 2, 6)
    controlLetter = Right(vkn, 1)
    if not IsAllNumeric(numericPart):
        errorMsg.value = "VKN'nin 2-7 karakterleri rakam olmalıdır."
        return _result
    if not (vba_like(controlLetter, "[A-Z]")):
        errorMsg.value = "VKN'nin son karakteri bir harf olmalıdır."
        return _result
    validFirstChars = "FALECDGOPU"
    if InStr(validFirstChars, firstChar) == 0:
        errorMsg.value = "VKN'nin ilk karakteri geçerli bir harf olmalıdır."
        return _result
    numValue = CLng(numericPart)
    validRange = False
    if firstChar == "F":
        if numValue >= 0 and numValue <= 699999:
            validRange = True
    elif firstChar == "E":
        if numValue >= 0 and numValue <= 999999:
            validRange = True
    elif firstChar == "A" or firstChar == "L":
        if numValue >= 700000 and numValue <= 799999:
            validRange = True
    elif firstChar == "C" or firstChar == "D" or firstChar == "G" or firstChar == "O" or firstChar == "P" or firstChar == "U":
        validRange = True
    else:
        errorMsg.value = "VKN'nin ilk karakteri geçerli bir tipte olmalıdır."
        return _result
    if not validRange:
        errorMsg.value = "VKN'nin sayısal kısmı geçerli aralıkta değil."
        return _result
    _result = True
    return _result

def ValidateAnguillaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    firstDigit = ""
    tin = Replace(Trim(tin), " ", "")
    if Len(tin) != 10:
        errorMsg.value = "Anguilla TIN'i 10 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Anguilla TIN'i tamamen rakamlardan oluşmalıdır."
        return _result
    firstDigit = Left(tin, 1)
    if firstDigit == "1" or firstDigit == "2":
        _result = True
    else:
        errorMsg.value = "Anguilla TIN'i bireyler için 1 ile, işletmeler için 2 ile başlamalıdır."
    return _result

def ValidateArgentinaCUIT(cuit, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    prefix = ""
    mainNumber = ""
    checkDigit = ""
    cuit = Replace(Replace(Replace(Trim(cuit), " ", ""), "-", ""), ".", "")
    if Len(cuit) != 11:
        errorMsg.value = "Arjantin CUIT'i 11 haneli olmalıdır."
        return _result
    if not IsAllNumeric(cuit):
        errorMsg.value = "Arjantin CUIT'i sadece rakamlardan oluşmalıdır."
        return _result
    prefix = Left(cuit, 2)
    mainNumber = Mid(cuit, 3, 8)
    checkDigit = Right(cuit, 1)
    if prefix == "20" or prefix == "23" or prefix == "24" or prefix == "27" or prefix == "30" or prefix == "33":
        _result = True
    else:
        errorMsg.value = "CUIT için geçersiz ön ek. Bireyler için 20, 23, 24 veya 27; tüzel kişiler için 30 veya 33 olmalıdır."
    return _result

def ValidateArmeniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if not vba_like(tin, "########"):
        errorMsg.value = "Ermenistan TIN'i tam olarak 8 rakamdan oluşmalıdır. Başka karakter içermemelidir."
        return _result
    _result = True
    return _result

def ValidateArubaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if Len(tin) != 8:
        errorMsg.value = "Aruba TIN'i 8 haneli olmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Aruba TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateAustraliaVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanVKN = ""
    cleanVKN = Replace(Replace(Trim(vkn), " ", ""), "-", "")
    if not IsAllNumeric(cleanVKN):
        errorMsg.value = "Avustralya VKN'si sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanVKN) == 8 or Len(cleanVKN) == 9 or Len(cleanVKN) == 11:
        _result = True
    else:
        errorMsg.value = "Avustralya VKN'si 8, 9 veya 11 haneli olmalıdır (boşluklar hariç)."
    return _result

def ValidateAustriaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    product = 0
    cleanTIN = Replace(Replace(Replace(Trim(tin), " ", ""), "-", ""), "/", "")
    if Len(cleanTIN) != 9:
        errorMsg.value = "Avusturya TIN'i 9 haneli olmalıdır (temizlendikten sonra)."
        _result = False
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Avusturya TIN'i sadece rakamlardan oluşmalıdır."
        _result = False
        return _result
    weightedSum = 0
    for i in vba_range(1, 8, 1):
        digit = CInt(Mid(cleanTIN, i, 1))
        if i % 2 != 0:
            product = digit * 1
        else:
            product = digit * 2
        if product > 9:
            product = product - 9
        weightedSum = weightedSum + product
    calculatedCheckDigit = (10 - (weightedSum % 10)) % 10
    checkDigit = CInt(Right(cleanTIN, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Avusturya TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        _result = False
        return _result
    _result = True
    return _result

def ValidateAzerbaijanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    i = 0
    for i in vba_range(1, Len(cleanTIN), 1):
        if not (vba_like(Mid(cleanTIN, i, 1), "[0-9A-Za-z]")):
            errorMsg.value = "Azerbaycan TIN'i sadece harf ve rakamlardan oluşmalıdır. Özel karakterler içeremez."
            return _result
    if Len(cleanTIN) == 10:
        if IsAllNumeric(cleanTIN):
            _result = True
        else:
            errorMsg.value = "10 haneli Azerbaycan TIN'i tamamen rakamlardan oluşmalıdır."
    elif Len(cleanTIN) == 7:
        _result = True
    else:
        errorMsg.value = "Azerbaycan TIN'i 10 haneli (tamamen rakam) veya 7 haneli (harf ve rakam) olmalıdır."
    return _result

def ValidateBahrainTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Replace(Trim(tin), " ", "")
    if Len(tin) != 15:
        errorMsg.value = "Bahreyn TIN'i 15 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Bahreyn TIN'i tamamen rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateBarbadosTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Replace(Trim(tin), " ", "")
    if Len(tin) != 13:
        errorMsg.value = "Barbados TIN must be 13 digits."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Barbados TIN must contain only digits."
        return _result
    if Left(tin, 1) != "1":
        errorMsg.value = "Barbados TIN must start with '1'."
        return _result
    _result = True
    return _result

def ValidateBelarusTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAlphanumeric(cleanTIN):
        errorMsg.value = "Belarus TIN'i sadece harf ve rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Belarus TIN'i 9 karakter olmalıdır."
        return _result
    _result = True
    return _result

def ValidateBelgiumTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    checkDigits = ""
    mainNumber = ""
    remainder1900 = 0
    remainder2000 = 0
    expectedCheckDigits1900 = ""
    expectedCheckDigits2000 = ""
    mainNumberLong = 0
    mainNumberMod97 = 0
    CENTURY_OFFSET_MOD97 = 68
    cleanTIN = Replace(Replace(Replace(Trim(tin), " ", ""), "-", ""), ".", "")
    if Len(cleanTIN) != 11:
        errorMsg.value = "Belçika TIN'i 11 haneli olmalıdır."
        _result = False
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Belçika TIN'i sadece rakamlardan oluşmalıdır."
        _result = False
        return _result
    mainNumber = Left(cleanTIN, 9)
    checkDigits = Right(cleanTIN, 2)
    mainNumberLong = CLng(mainNumber)
    mainNumberMod97 = mainNumberLong % 97
    remainder1900 = mainNumberMod97
    if remainder1900 == 0:
        expectedCheckDigits1900 = "97"
    else:
        expectedCheckDigits1900 = CStr(97 - remainder1900)
        if Len(expectedCheckDigits1900) == 1:
            expectedCheckDigits1900 = "0" + expectedCheckDigits1900
    remainder2000 = (mainNumberMod97 + CENTURY_OFFSET_MOD97) % 97
    if remainder2000 == 0:
        expectedCheckDigits2000 = "97"
    else:
        expectedCheckDigits2000 = CStr(97 - remainder2000)
        if Len(expectedCheckDigits2000) == 1:
            expectedCheckDigits2000 = "0" + expectedCheckDigits2000
    if (checkDigits == expectedCheckDigits1900) or (checkDigits == expectedCheckDigits2000):
        _result = True
    else:
        errorMsg.value = "Geçersiz Belçika TIN. Kontrol basamakları uyuşmuyor. Hesaplanan 1900'ler için: " + expectedCheckDigits1900 + ", 2000'ler için: " + expectedCheckDigits2000 + ", Girilen: " + checkDigits
        _result = False
    return _result

def ValidateBelizeTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Replace(Trim(tin), " ", "")
    if Len(tin) != 6:
        errorMsg.value = "Belize TIN 6 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Belize TIN tamamen rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateBhutanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) != 8:
        errorMsg.value = "Butan TIN'i 8 karakter olmalıdır."
        return _result
    if not (vba_like(Left(cleanTIN, 3), "[A-Z][A-Z][A-Z]") and vba_like(Right(cleanTIN, 5), "#####")):
        errorMsg.value = "Butan TIN'i AAA##### formatında olmalıdır (harf + harf + harf + rakam + rakam + rakam + rakam + rakam)."
        return _result
    _result = True
    return _result

def ValidateBoliviaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Bolivya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 7 and Len(cleanTIN) != 10:
        errorMsg.value = "Bolivya TIN'i 7 veya 10 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateBosniaHerzegovinaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Bosna Hersek TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 12:
        errorMsg.value = "Bosna Hersek TIN'i 12 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateBotswanaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if not vba_like(Left(cleanTIN, 1), "[A-Z]"):
        errorMsg.value = "Botsvana TIN'i bir harf ile başlamalıdır (A-Z)."
        return _result
    if Len(cleanTIN) == 10:
        numericPart = Right(cleanTIN, 9)
        if IsAllNumeric(numericPart):
            _result = True
            return _result
        else:
            errorMsg.value = "Botsvana TIN'i bir harf ile başlamalı ve ardından 9 rakam gelmelidir."
            return _result
    elif Len(cleanTIN) == 11:
        numericPart = Right(cleanTIN, 10)
        if IsAllNumeric(numericPart):
            _result = True
            return _result
        else:
            errorMsg.value = "Botsvana TIN'i bir harf ile başlamalı ve ardından 10 rakam gelmelidir."
            return _result
    else:
        errorMsg.value = "Botsvana TIN'i bir harf ile başlamalı ve ardından 9 veya 10 rakam gelmelidir."
        return _result
    return _result

def ValidateBrazilCPF_CNPJ(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    currentChar = ""
    errorMsg.value = ""
    if Len(tin) != 11 and Len(tin) != 14:
        errorMsg.value = "Brezilya TIN'i 11 haneli (CPF) veya 14 haneli (CNPJ) olmalıdır."
        _result = False
        return _result
    for i in vba_range(1, Len(tin), 1):
        currentChar = Mid(tin, i, 1)
        if not (currentChar >= "0" and currentChar <= "9"):
            errorMsg.value = "Brezilya TIN'i sadece rakamlardan oluşmalıdır."
            _result = False
            return _result
    _result = True
    return _result

def ValidateCPF(cleanCPF, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    sum = 0
    firstCheckDigit = 0
    secondCheckDigit = 0
    if cleanCPF == StringFunc(Len(cleanCPF), Mid(cleanCPF, 1, 1)):
        errorMsg.value = "Geçersiz CPF: Tüm rakamlar aynı."
        _result = False
        return _result
    sum = 0
    for i in vba_range(1, 9, 1):
        sum = sum + CInt(Mid(cleanCPF, i, 1)) * (10 - i)
    firstCheckDigit = (sum * 10) % 11
    if firstCheckDigit == 10:
        firstCheckDigit = 0
    if firstCheckDigit != CInt(Mid(cleanCPF, 10, 1)):
        errorMsg.value = "Geçersiz CPF: İlk kontrol basamağı uyuşmuyor."
        _result = False
        return _result
    sum = 0
    for i in vba_range(1, 10, 1):
        sum = sum + CInt(Mid(cleanCPF, i, 1)) * (11 - i)
    secondCheckDigit = (sum * 10) % 11
    if secondCheckDigit == 10:
        secondCheckDigit = 0
    if secondCheckDigit != CInt(Mid(cleanCPF, 11, 1)):
        errorMsg.value = "Geçersiz CPF: İkinci kontrol basamağı uyuşmuyor."
        _result = False
        return _result
    _result = True
    return _result

def ValidateBruneiTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    hyphenCount = 0
    i = 0
    char = ""
    cleanTIN = Trim(tin)
    hyphenCount = 0
    for i in vba_range(1, Len(cleanTIN), 1):
        char = Mid(cleanTIN, i, 1)
        if char == "-":
            hyphenCount = hyphenCount + 1
    if hyphenCount > 1:
        errorMsg.value = "Brunei TIN'i en fazla 1 adet tire (-) karakteri içerebilir."
        return _result
    if Len(cleanTIN) < 9 or Len(cleanTIN) > 11:
        errorMsg.value = "Brunei TIN'i 9, 10 veya 11 karakterden oluşmalıdır."
        return _result
    for i in vba_range(1, Len(cleanTIN), 1):
        char = Mid(cleanTIN, i, 1)
        if char != "-":
            if not (vba_like(char, "[A-Z]") or vba_like(char, "[0-9]")):
                errorMsg.value = "Brunei TIN'i harf, rakam veya tire (-) karakteri içermelidir. Boşluk veya özel karakterler içeremez."
                return _result
    _result = True
    return _result

def ValidateBulgariaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Bulgaristan TIN'i sadece rakamlardan oluşmalıdır (boşluk veya tire olmadan)."
        return _result
    if Len(cleanTIN) != 10:
        errorMsg.value = "Bulgaristan TIN'i 10 haneli olmalıdır."
        return _result
    weights = Array(2, 4, 8, 5, 10, 9, 7, 3, 6)
    weightedSum = 0
    for i in vba_range(1, 9, 1):
        digit = CInt(Mid(cleanTIN, i, 1))
        weightedSum = weightedSum + digit * weights(i - 1)
    calculatedCheckDigit = weightedSum % 11
    if calculatedCheckDigit == 10:
        calculatedCheckDigit = 0
    checkDigit = CInt(Right(cleanTIN, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Bulgaristan TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateBurkinaFasoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) != 9:
        errorMsg.value = "Burkina Faso TIN'i 9 karakter olmalıdır (8 rakam ve ardından bir harf)."
        return _result
    if not vba_like(Right(cleanTIN, 1), "[A-Z]"):
        errorMsg.value = "Burkina Faso TIN'inin son karakteri harf olmalıdır (A-Z)."
        return _result
    numericPart = Left(cleanTIN, 8)
    if not IsAllNumeric(numericPart):
        errorMsg.value = "Burkina Faso TIN'inin ilk 8 karakteri rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateBurundiTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Burundi TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10:
        errorMsg.value = "Burundi TIN'i 10 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateCambodiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    hyphenCount = 0
    i = 0
    char = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) != 13 and Len(cleanTIN) != 14:
        errorMsg.value = "Kamboçya TIN'i tam olarak 13 veya 14 karakter olmalıdır."
        return _result
    hyphenCount = 0
    for i in vba_range(1, Len(cleanTIN), 1):
        char = Mid(cleanTIN, i, 1)
        if char == "-":
            hyphenCount = hyphenCount + 1
    if hyphenCount != 1:
        errorMsg.value = "Kamboçya TIN'i tam olarak 1 adet tire (-) içermelidir."
        return _result
    if not vba_like(Left(cleanTIN, 1), "[A-Z]"):
        errorMsg.value = "Kamboçya TIN'inin ilk karakteri harf (A-Z) olmalıdır."
        return _result
    for i in vba_range(2, Len(cleanTIN), 1):
        char = Mid(cleanTIN, i, 1)
        if char != "-":
            if not (vba_like(char, "[0-9]")):
                errorMsg.value = "Kamboçya TIN'inin sonraki karakterleri sadece rakamlar veya tire(-) olmalıdır."
                return _result
    _result = True
    return _result

def ValidateCameroonNIU(niu, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanNIU = ""
    numericPart = ""
    cleanNIU = Replace(Trim(niu), " ", "")
    cleanNIU = UCase(cleanNIU)
    if Len(cleanNIU) != 14:
        errorMsg.value = "Kamerun NIU'su 14 karakter olmalıdır."
        return _result
    if not (vba_like(Left(cleanNIU, 1), "[A-Z]") and vba_like(Right(cleanNIU, 1), "[A-Z]")):
        errorMsg.value = "Kamerun NIU'su bir harf ile başlamalı ve bir harf ile bitmelidir (A-Z)."
        return _result
    numericPart = Mid(cleanNIU, 2, 12)
    if not IsAllNumeric(numericPart):
        errorMsg.value = "Kamerun NIU'sunun 2-13 arasındaki karakterleri sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateCanadaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if Len(cleanTIN) == 9 and IsAllNumeric(cleanTIN):
        _result = True
        return _result
    if Len(cleanTIN) == 9 and IsAllNumeric(cleanTIN):
        _result = True
        return _result
    if Len(cleanTIN) == 9 and Left(cleanTIN, 1) == "T" and IsAllNumeric(Right(cleanTIN, 8)):
        _result = True
        return _result
    errorMsg.value = "Kanada TIN'i geçersiz. 9 haneli SIN/BN veya 'T' ile başlayan 8 haneli Trust Hesap Numarası olmalıdır."
    return _result

def ValidateCentralAfricanRepublicTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) != 8:
        errorMsg.value = "Orta Afrika Cumhuriyeti TIN'i 8 karakter olmalıdır."
        return _result
    if not vba_like(Right(cleanTIN, 1), "[A-Z]"):
        errorMsg.value = "Orta Afrika Cumhuriyeti TIN'inin son karakteri bir harf olmalıdır (A-Z)."
        return _result
    numericPart = Left(cleanTIN, 7)
    if not IsAllNumeric(numericPart):
        errorMsg.value = "Orta Afrika Cumhuriyeti TIN'inin ilk 7 karakteri rakam olmalıdır."
        return _result
    _result = True
    return _result

def ValidateChileTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    serialNumber = ""
    checkDigit = ""
    cleanTIN = ""
    i = 0
    multiplier = 0
    totalSum = 0
    remainder = 0
    calculatedCheckDigit = ""
    currentDigit = 0
    cleanTIN = Replace(Replace(Replace(Trim(tin), ".", ""), "-", ""), " ", "")
    if Len(cleanTIN) < 8 or Len(cleanTIN) > 9:
        errorMsg.value = "Şili TIN toplam 8 veya 9 karakter olmalıdır."
        return _result
    serialNumber = Left(cleanTIN, Len(cleanTIN) - 1)
    checkDigit = Right(cleanTIN, 1)
    if not IsAllNumeric(serialNumber):
        errorMsg.value = "TIN'in seri numarası kısmı (kontrol rakamı hariç) sadece rakamlardan oluşmalıdır."
        return _result
    if Len(serialNumber) < 7 or Len(serialNumber) > 8:
        errorMsg.value = "TIN'in seri numarası kısmı 7 veya 8 rakamdan oluşmalıdır."
        return _result
    if not (vba_like(UCase(checkDigit), "[0-9K]")):
        errorMsg.value = "TIN'in kontrol rakamı 0-9 veya 'K' olmalıdır."
        return _result
    totalSum = 0
    multiplier = 2
    for i in vba_range(Len(serialNumber), 1, -1):
        currentDigit = CInt(Mid(serialNumber, i, 1))
        totalSum = totalSum + currentDigit * multiplier
        multiplier = multiplier + 1
        if multiplier > 7:
            multiplier = 2
    remainder = 11 - (totalSum % 11)
    if remainder == 11:
        calculatedCheckDigit = "0"
    elif remainder == 10:
        calculatedCheckDigit = "K"
    else:
        calculatedCheckDigit = CStr(remainder)
    if UCase(checkDigit) == calculatedCheckDigit:
        _result = True
    else:
        errorMsg.value = "Geçersiz Şili TIN: Beklenen kontrol rakamı '" + calculatedCheckDigit + "', ancak '" + UCase(checkDigit) + "' girildi."
    return _result

def ValidateChinaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    charCode = 0
    if Len(tin) == 15 or Len(tin) == 18:
        for i in vba_range(1, Len(tin), 1):
            charCode = Asc(Mid(tin, i, 1))
            if not ((charCode >= 48 and charCode <= 57) or (charCode >= 65 and charCode <= 90) or (charCode >= 97 and charCode <= 122)):
                errorMsg.value = "Çin TIN'i sadece rakamlar ve harflerden oluşmalıdır. Boşluk veya özel karakter içeremez."
                return _result
        _result = True
    else:
        errorMsg.value = "Çin TIN'i 15 veya 18 karakter uzunluğunda olmalıdır."
    return _result

def ValidateColombiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanedTIN = ""
    tinLength = 0
    mainNumber = ""
    checkDigit = ""
    mainNumberValue = 0
    isLegalEntity = False
    isIndividual = False
    i = 0
    for i in vba_range(1, Len(tin), 1):
        if not vba_like(Mid(tin, i, 1), "#"):
            errorMsg.value = "Kolombiya TIN'i sadece rakamlardan oluşmalıdır. Harf, boşluk veya özel karakter içeremez."
            return _result
    cleanedTIN = tin
    tinLength = Len(cleanedTIN)
    if tinLength < 2:
        errorMsg.value = "Kolombiya TIN'i en az 2 haneli olmalıdır (1 hane + 1 doğrulama basamağı)."
        return _result
    checkDigit = Right(cleanedTIN, 1)
    mainNumber = Left(cleanedTIN, tinLength - 1)
    mainNumberValue = Val(mainNumber)
    if Len(mainNumber) < 1 or Len(mainNumber) > 13:
        errorMsg.value = "Kolombiya TIN'inin ana numarası 1 ile 13 hane arasında olmalıdır."
        return _result
    if Len(mainNumber) == 9:
        if mainNumberValue >= 800000000 and mainNumberValue <= 899999999:
            isLegalEntity = True
        elif mainNumberValue >= 900000000:
            isLegalEntity = True
    if mainNumberValue >= 1 and mainNumberValue <= 99999999:
        isIndividual = True
    elif mainNumberValue >= 700000001 and mainNumberValue <= 799999999:
        isIndividual = True
    elif mainNumberValue >= 600000001 and mainNumberValue <= 799999999:
        isIndividual = True
    elif mainNumberValue >= 1000000000 and Len(mainNumber) <= 13:
        isIndividual = True
    if not isLegalEntity and not isIndividual:
        errorMsg.value = "Kolombiya TIN'i geçerli bir aralıkta değil."
        return _result
    _result = True
    return _result
    return _result

def ValidateCookIslandsTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Replace(Trim(tin), " ", "")
    if Len(tin) != 5:
        errorMsg.value = "Cook Adaları TIN tam olarak 5 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Cook Adaları TIN sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateCostaRicaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Replace(Trim(tin), "-", ""), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "TIN sadece rakamlardan oluşmalıdır."
        return _result
    tinLength = 0
    tinLength = Len(cleanTIN)
    firstDigit = ""
    firstThreeDigits = ""
    firstFourDigits = ""
    isValid = False
    isValid = False
    if tinLength == 9:
        firstDigit = Left(cleanTIN, 1)
        if vba_like(firstDigit, "[1-7]"):
            isValid = True
        else:
            errorMsg.value = "Geçersiz bireysel TIN formatı."
            return _result
    elif tinLength == 10:
        firstDigit = Left(cleanTIN, 1)
        firstFourDigits = Left(cleanTIN, 4)
        if firstFourDigits == "3120":
            isValid = True
        elif firstFourDigits == "3130":
            isValid = True
        elif vba_like(firstDigit, "[2-3]"):
            isValid = True
        else:
            errorMsg.value = "Geçersiz kurumsal TIN veya NITE formatı."
            return _result
    elif tinLength == 11 or tinLength == 12:
        isValid = True
    else:
        errorMsg.value = "TIN uzunluğu geçerli değil."
        return _result
    if isValid:
        _result = True
    return _result

def ValidateCroatiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    digits = VbaArray(1, 11, 0)
    sum = 0
    subtotal = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    if Len(tin) != 11:
        errorMsg.value = "Hırvatistan TIN'i tam olarak 11 rakamdan oluşmalıdır."
        return _result
    for i in vba_range(1, 11, 1):
        if not vba_like(Mid(tin, i, 1), "#"):
            errorMsg.value = "Hırvatistan TIN'i sadece rakamlardan oluşmalıdır."
            return _result
        digits[i] = CInt(Mid(tin, i, 1))
    sum = 10
    for i in vba_range(1, 10, 1):
        sum = (digits(i) + sum) % 10
        if sum == 0:
            sum = 10
        sum = (sum * 2) % 11
    if sum == 1:
        calculatedCheckDigit = 0
    else:
        calculatedCheckDigit = 11 - sum
    checkDigit = digits(11)
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Hırvatistan TIN'i. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateCuracaoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if not IsAllNumeric(tin):
        errorMsg.value = "Curaçao TIN sadece rakamlardan oluşmalıdır ve özel karakter içeremez."
        return _result
    if Len(tin) != 9:
        errorMsg.value = "Curaçao TIN tam olarak 9 rakamdan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateCyprusTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    c = VbaArray(1, 8, 0)
    c9 = ""
    i = 0
    even_sum = 0
    odd_sum = 0
    total_sum = 0
    remainder = 0
    check_digit_ascii = 0
    calculated_check_digit = ""
    conversion_table = VbaArray(0, 9, 0)
    firstDigit = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) != 9:
        errorMsg.value = "Kıbrıs TIN tam olarak 9 karakterden oluşmalıdır (8 rakam + 1 büyük harf)."
        return _result
    for i in vba_range(1, 8, 1):
        if not (vba_like(Mid(cleanTIN, i, 1), "[0-9]")):
            errorMsg.value = "TIN'in ilk 8 karakteri sadece rakamlardan oluşmalıdır."
            return _result
        c[i] = CInt(Mid(cleanTIN, i, 1))
    c9 = Mid(cleanTIN, 9, 1)
    if not (vba_like(c9, "[A-Z]")):
        errorMsg.value = "TIN'in son karakteri büyük harf olmalıdır (A-Z)."
        return _result
    firstDigit = Mid(cleanTIN, 1, 1)
    if firstDigit != "0" and firstDigit != "9":
        errorMsg.value = "Bireyler için TIN'in ilk rakamı 0 veya 9 olmalıdır."
        return _result
    conversion_table[0] = 1
    conversion_table[1] = 0
    conversion_table[2] = 5
    conversion_table[3] = 7
    conversion_table[4] = 9
    conversion_table[5] = 13
    conversion_table[6] = 15
    conversion_table[7] = 17
    conversion_table[8] = 19
    conversion_table[9] = 21
    even_sum = c(2) + c(4) + c(6) + c(8)
    odd_sum = conversion_table(c(1)) + conversion_table(c(3)) + conversion_table(c(5)) + conversion_table(c(7))
    total_sum = even_sum + odd_sum
    remainder = total_sum % 26
    check_digit_ascii = remainder + 65
    calculated_check_digit = Chr(check_digit_ascii)
    if calculated_check_digit != c9:
        errorMsg.value = "Geçersiz TIN: Beklenen kontrol harfi '" + calculated_check_digit + "', ancak '" + c9 + "' girildi."
        return _result
    _result = True
    return _result

def ValidateCzechiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "/", "")
    if Len(cleanTIN) != 9 and Len(cleanTIN) != 10:
        errorMsg.value = "Çekya TIN'i 9 veya 10 haneli olmalıdır."
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Çekya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateDemocraticRepublicOfCongoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if not IsAlphanumeric(cleanTIN):
        errorMsg.value = "Kongo Demokratik Cumhuriyeti TIN'i sadece harf ve rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) < 7 or Len(cleanTIN) > 9:
        errorMsg.value = "Kongo Demokratik Cumhuriyeti TIN'i 7, 8 veya 9 karakter olmalıdır."
        return _result
    _result = True
    return _result

def ValidateDenmarkTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if Len(cleanTIN) != 10:
        errorMsg.value = "Danimarka CPR numarası 10 haneli olmalıdır."
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Danimarka CPR numarası sadece rakamlardan oluşmalıdır."
        return _result
    weights = Array(4, 3, 2, 7, 6, 5, 4, 3, 2)
    weightedSum = 0
    for i in vba_range(1, 9, 1):
        digit = CInt(Mid(cleanTIN, i, 1))
        weightedSum = weightedSum + digit * weights(i - 1)
    calculatedCheckDigit = weightedSum % 11
    if calculatedCheckDigit == 1:
        errorMsg.value = "Geçersiz Danimarka TIN. Modulo sonucu 1 olamaz."
        return _result
    elif calculatedCheckDigit == 0:
        calculatedCheckDigit = 0
    else:
        calculatedCheckDigit = 11 - calculatedCheckDigit
    checkDigit = CInt(Right(cleanTIN, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Danimarka TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateDominicaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinLength = 0
    mainNumber = ""
    checkDigit = ""
    tin = Replace(Trim(tin), " ", "")
    tinLength = Len(tin)
    if tinLength == 6:
        mainNumber = tin
        if IsAllNumeric(mainNumber):
            _result = True
        else:
            errorMsg.value = "Dominika TIN sadece rakamlardan oluşmalıdır."
    elif tinLength == 7:
        mainNumber = Left(tin, 6)
        checkDigit = Right(tin, 1)
        if IsAllNumeric(mainNumber) and IsNumeric(checkDigit):
            _result = True
        else:
            errorMsg.value = "Dominika TIN sadece rakamlardan oluşmalıdır."
    else:
        errorMsg.value = "Dominika TIN 6 veya 7 rakamdan oluşmalıdır."
    return _result

def ValidateEcuadorRUC(ruc, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    ruc = Trim(ruc)
    if Len(ruc) != 13:
        errorMsg.value = "Ekvador RUC'su 13 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(ruc):
        errorMsg.value = "Ekvador RUC'su sadece rakamlardan oluşmalıdır. Özel karakter veya harf içeremez."
        return _result
    if Right(ruc, 3) != "001":
        errorMsg.value = "Ekvador RUC'sunun son üç rakamı '001' olmalıdır."
        return _result
    _result = True
    return _result

def ValidateEgyptTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Mısır TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Mısır TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateEstoniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    cleanTIN = ""
    for i in vba_range(1, Len(tin), 1):
        if IsNumeric(Mid(tin, i, 1)):
            cleanTIN = cleanTIN + Mid(tin, i, 1)
    if Len(cleanTIN) != 11:
        errorMsg.value = "Estonya TIN'i 11 haneli olmalıdır."
        return _result
    weights = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 1)
    calculatedCheckDigit = CalculateCheckDigit(cleanTIN, weights)
    if calculatedCheckDigit == 10:
        weights = Array(3, 4, 5, 6, 7, 8, 9, 1, 2, 3)
        calculatedCheckDigit = CalculateCheckDigit(cleanTIN, weights)
        if calculatedCheckDigit == 10:
            calculatedCheckDigit = 0
    checkDigit = CInt(Right(cleanTIN, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Estonya TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    firstDigit = 0
    year = 0
    month = 0
    day = 0
    serial = 0
    firstDigit = CInt(Left(cleanTIN, 1))
    if firstDigit < 1 or firstDigit > 6:
        errorMsg.value = "Estonya TIN'inin ilk hanesi 1 ile 6 arasında olmalıdır."
        return _result
    year = CInt(Mid(cleanTIN, 2, 2))
    month = CInt(Mid(cleanTIN, 4, 2))
    day = CInt(Mid(cleanTIN, 6, 2))
    serial = CInt(Mid(cleanTIN, 8, 3))
    if month < 1 or month > 12:
        errorMsg.value = "Geçersiz ay."
        return _result
    if day < 1 or day > 31:
        errorMsg.value = "Geçersiz gün."
        return _result
    if serial < 1 or serial > 710:
        errorMsg.value = "Seri numarası 001 ile 710 arasında olmalıdır."
        return _result
    _result = True
    return _result

def ValidateFaroeIslandsTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinClean = ""
    datePart = ""
    tinDay = 0
    tinMonth = 0
    tinYear = 0
    genderDigit = 0
    testDate = None
    tinClean = Replace(Replace(Trim(tin), "-", ""), " ", "")
    if Len(tinClean) != 9:
        errorMsg.value = "Faroe Adaları TIN'i 9 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tinClean):
        errorMsg.value = "Faroe Adaları TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    datePart = Left(tinClean, 6)
    tinDay = CInt(Left(datePart, 2))
    tinMonth = CInt(Mid(datePart, 3, 2))
    tinYear = CInt(Right(datePart, 2))
    if tinYear >= 50:
        tinYear = 1900 + tinYear
    else:
        tinYear = 2000 + tinYear
    testDate = DateSerial(tinYear, tinMonth, tinDay)
    if tinDay != VBA.day(testDate) or tinMonth != VBA.month(testDate) or tinYear != VBA.year(testDate):
        errorMsg.value = "Faroe Adaları TIN'inin ilk 6 rakamı geçerli bir tarih olmalıdır."
        return _result
    genderDigit = CInt(Right(tinClean, 1))
    _result = True
    return _result
    errorMsg.value = "Faroe Adaları TIN'inin ilk 6 rakamı geçerli bir tarih olmalıdır."
    return _result

def ValidateFinlandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    checkDigit = ""
    mainNumber = ""
    calculatedCheckDigit = ""
    moduloResult = 0
    checkChars = ""
    checkTable = None
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) != 11:
        errorMsg.value = "Finlandiya TIN'i 11 karakter olmalıdır."
        return _result
    if Mid(cleanTIN, 7, 1) != "-" and Mid(cleanTIN, 7, 1) != "+" and Mid(cleanTIN, 7, 1) != "A":
        errorMsg.value = "Finlandiya TIN'i 7. pozisyonda (+, - veya A) ayırıcı karakter içermelidir."
        return _result
    checkDigit = Right(cleanTIN, 1)
    mainNumber = Left(cleanTIN, 6) + Mid(cleanTIN, 8, 3)
    if not IsAllNumeric(mainNumber):
        errorMsg.value = "Finlandiya TIN'inin ilk 6 ve son 3 hanesi sayısal olmalıdır."
        return _result
    moduloResult = CLng(mainNumber) % 31
    checkTable = CreateObject("Scripting.Dictionary")
    checkTable.Add(0, "0")
    checkTable.Add(1, "1")
    checkTable.Add(2, "2")
    checkTable.Add(3, "3")
    checkTable.Add(4, "4")
    checkTable.Add(5, "5")
    checkTable.Add(6, "6")
    checkTable.Add(7, "7")
    checkTable.Add(8, "8")
    checkTable.Add(9, "9")
    checkTable.Add(10, "A")
    checkTable.Add(11, "B")
    checkTable.Add(12, "C")
    checkTable.Add(13, "D")
    checkTable.Add(14, "E")
    checkTable.Add(15, "F")
    checkTable.Add(16, "H")
    checkTable.Add(17, "J")
    checkTable.Add(18, "K")
    checkTable.Add(19, "L")
    checkTable.Add(20, "M")
    checkTable.Add(21, "N")
    checkTable.Add(22, "P")
    checkTable.Add(23, "R")
    checkTable.Add(24, "S")
    checkTable.Add(25, "T")
    checkTable.Add(26, "U")
    checkTable.Add(27, "V")
    checkTable.Add(28, "W")
    checkTable.Add(29, "X")
    checkTable.Add(30, "Y")
    calculatedCheckDigit = checkTable(moduloResult)
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Finlandiya TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateFranceVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if (Len(vkn) == 13 or Len(vkn) == 9) and IsAllNumeric(vkn):
        _result = True
    else:
        errorMsg.value = "Fransa VKN'si 9 veya 13 karakter olmalı ve tamamen rakamlardan oluşmalıdır."
    return _result

def ValidateGeorgiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if not vba_like(tin, "#######"):
        errorMsg.value = "Gürcistan TIN'i tam olarak 7 rakamdan oluşmalıdır. Başka karakter içermemelidir."
        return _result
    _result = True
    return _result

def ValidateGermanyVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(vkn) != 11:
        errorMsg.value = "Almanya VKN'si 11 karakter olmalıdır."
        return _result
    if not IsAllNumeric(vkn):
        errorMsg.value = "Almanya VKN'si tamamen rakamlardan oluşmalıdır."
        return _result
    if Left(vkn, 1) == "0":
        errorMsg.value = "Almanya VKN'si 0 ile başlayamaz."
        return _result
    _result = True
    return _result

def ValidateGhanaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinClean = ""
    tinLength = 0
    prefix = ""
    mainPart = ""
    validPrefixes = None
    countryCode = ""
    numberPart = ""
    checksum = ""
    i = 0
    tinClean = Replace(Replace(Trim(tin), " ", ""), "-", "")
    tinLength = Len(tinClean)
    if tinLength == 11:
        prefix = Left(tinClean, 3)
        mainPart = Mid(tinClean, 4, 8)
        validPrefixes = Array("P00", "C00", "G00", "Q00", "V00")
        if IsError(Application.Match(prefix, validPrefixes, 0)):
            errorMsg.value = "GRA TIN geçersiz prefix içeriyor."
            _result = False
            return _result
        if not IsAlphanumeric(mainPart):
            errorMsg.value = "GRA TIN'in ana kısmı alfanümerik olmalıdır."
            _result = False
            return _result
        _result = True
        return _result
    elif tinLength == 13 or tinLength == 15:
        if tinLength == 15:
            tinClean = Replace(tin, "-", "")
            tinClean = Replace(tinClean, " ", "")
            if Len(tinClean) != 13:
                errorMsg.value = "NIA Ghanacard PIN formatı hatalı. Doğru format: XXX-XXXXXXXXX-X"
                _result = False
                return _result
        countryCode = Left(tinClean, 3)
        numberPart = Mid(tinClean, 4, Len(tinClean) - 4)
        checksum = Right(tinClean, 1)
        if Len(countryCode) != 3 or not IsAllLetters(countryCode):
            errorMsg.value = "NIA Ghanacard PIN'in ülke kodu geçersiz. 3 harfli ISO kodu olmalıdır."
            _result = False
            return _result
        if not IsAllNumeric(numberPart):
            errorMsg.value = "NIA Ghanacard PIN'in numara kısmı sadece rakamlardan oluşmalıdır."
            _result = False
            return _result
        if Len(numberPart) != 9:
            errorMsg.value = "NIA Ghanacard PIN'in numara kısmı 9 haneli olmalıdır."
            _result = False
            return _result
        if Len(checksum) != 1 or not IsAlphanumeric(checksum):
            errorMsg.value = "NIA Ghanacard PIN'in checksum karakteri geçersiz. Tek bir rakam veya harf olmalıdır."
            _result = False
            return _result
        _result = True
        return _result
    else:
        errorMsg.value = "Gana TIN'in uzunluğu geçersizdir. GRA TIN 11 karakter veya NIA Ghanacard PIN 15 karakter olmalıdır."
        _result = False
    return _result

def ValidateGibraltarTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Trim(tin)
    if tin == "":
        errorMsg.value = "Gibraltar VKN boş olamaz."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Gibraltar VKN sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateGreeceTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if not vba_like(tin, "#########"):
        errorMsg.value = "Yunanistan TIN'i tam olarak 9 rakamdan oluşmalıdır. Başka karakter içermemelidir."
        return _result
    _result = True
    return _result

def ValidateGreenlandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if Len(cleanTIN) == 10:
        if IsAllNumeric(cleanTIN):
            _result = True
            return _result
        else:
            errorMsg.value = "Gerçek kişiler için Grönland TIN'i 10 haneli ve sadece rakamlardan oluşmalıdır."
            return _result
    if Len(cleanTIN) == 8:
        if IsAllNumeric(cleanTIN):
            _result = True
            return _result
        else:
            errorMsg.value = "Tüzel kişiler için Grönland TIN'i 8 haneli ve sadece rakamlardan oluşmalıdır."
            return _result
    errorMsg.value = "Grönland TIN'i geçersiz. 10 haneli (gerçek kişiler) veya 8 haneli (tüzel kişiler) olmalıdır."
    return _result

def ValidateGrenadaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Trim(tin)
    if Len(tin) != 6:
        errorMsg.value = "Grenada VKN'si tam olarak 6 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Grenada VKN'si sadece rakamlardan oluşmalıdır. Harf veya özel karakter içeremez."
        return _result
    _result = True
    return _result

def ValidateGuatemalaNIT(nit, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanNIT = ""
    cleanNIT = Replace(Trim(nit), " ", "")
    if not IsAllNumeric(cleanNIT):
        errorMsg.value = "Guatemala NIT'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanNIT) != 7 and Len(cleanNIT) != 8:
        errorMsg.value = "Guatemala NIT'i 7 veya 8 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateHaitiTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Haiti TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10:
        errorMsg.value = "Haiti TIN'i 10 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateHKID(hkid, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    prefix = ""
    numbers = ""
    checkDigit = ""
    prefixLength = 0
    totalLength = 0
    totalLength = Len(hkid)
    if totalLength != 8 and totalLength != 9:
        errorMsg.value = "HKID numarası 8 veya 9 karakter uzunluğunda olmalıdır."
        return _result
    if IsLetter(Mid(hkid, 2, 1)):
        prefix = Left(hkid, 2)
        numbers = Mid(hkid, 3, 6)
        checkDigit = Right(hkid, 1)
    else:
        prefix = Left(hkid, 1)
        numbers = Mid(hkid, 2, 6)
        checkDigit = Right(hkid, 1)
    if not IsAllLetters(prefix):
        errorMsg.value = "HKID numarasının başındaki karakter(ler) büyük harf olmalıdır."
        return _result
    if not IsAllNumeric(numbers):
        errorMsg.value = "HKID numarasının ortasındaki 6 karakter rakamlardan oluşmalıdır."
        return _result
    if not (IsNumeric(checkDigit) or checkDigit == "A" or checkDigit == "a"):
        errorMsg.value = "HKID numarasının son karakteri 0-9 veya 'A' olmalıdır."
        return _result
    _result = True
    return _result

def ValidateBRNumber(brNumber, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(brNumber) != 8:
        errorMsg.value = "BR numarası 8 rakamdan oluşmalıdır."
        _result = False
        return _result
    if not IsAllNumeric(brNumber):
        errorMsg.value = "BR numarası sadece rakamlardan oluşmalıdır."
        _result = False
        return _result
    _result = True
    return _result

def ValidateHongKongTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinClean = ""
    tinClean = Replace(Replace(Replace(Trim(tin), " ", ""), "(", ""), ")", "")
    if ValidateHKID(tinClean, errorMsg):
        _result = True
        return _result
    if ValidateBRNumber(tinClean, errorMsg):
        _result = True
        return _result
    errorMsg.value = "Hong Kong TIN formatı geçersiz. Bireyler için HKID numarası formatına veya kurumlar için BR numarası formatına uygun olmalıdır."
    _result = False
    return _result

def ValidateHungaryTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    digits = VbaArray(1, 10, 0)
    weights = VbaArray(1, 9, 0)
    sum = 0
    calculatedCheckDigit = 0
    checkDigit = 0
    if Len(tin) != 10:
        errorMsg.value = "Macaristan TIN'i tam olarak 10 rakamdan oluşmalıdır."
        return _result
    for i in vba_range(1, 10, 1):
        if not vba_like(Mid(tin, i, 1), "#"):
            errorMsg.value = "Macaristan TIN'i sadece rakamlardan oluşmalıdır."
            return _result
        digits[i] = CInt(Mid(tin, i, 1))
    if digits(1) != 8:
        errorMsg.value = "Macaristan TIN'inin ilk hanesi 8 olmalıdır."
        return _result
    for i in vba_range(1, 9, 1):
        weights[i] = i
    sum = 0
    for i in vba_range(1, 9, 1):
        sum = sum + digits(i) * weights(i)
    calculatedCheckDigit = sum % 11
    checkDigit = digits(10)
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Macaristan TIN'i. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateIcelandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinClean = ""
    tinDay = 0
    tinMonth = 0
    tinYear = 0
    tinCentury = 0
    birthdate = None
    testDate = None
    centuryIndicator = ""
    tinClean = Replace(Replace(Trim(tin), "-", ""), " ", "")
    if Len(tinClean) != 10:
        errorMsg.value = "İzlanda VKN'si 10 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tinClean):
        errorMsg.value = "İzlanda VKN'si sadece rakamlardan oluşmalıdır."
        return _result
    tinDay = CInt(Mid(tinClean, 1, 2))
    tinMonth = CInt(Mid(tinClean, 3, 2))
    tinYear = CInt(Mid(tinClean, 5, 2))
    centuryIndicator = Right(tinClean, 1)
    if centuryIndicator == "9":
        tinCentury = 1900
    elif centuryIndicator == "0":
        tinCentury = 2000
    elif centuryIndicator == "1":
        tinCentury = 2100
    else:
        errorMsg.value = "İzlanda VKN'sinin 10. rakamı yüzyılı belirtmelidir (9, 0 veya 1)."
        return _result
    tinYear = tinCentury + tinYear
    birthdate = DateSerial(tinYear, tinMonth, tinDay)
    if day(birthdate) != tinDay or month(birthdate) != tinMonth or year(birthdate) != tinYear:
        errorMsg.value = "İzlanda VKN'sinin ilk 6 rakamı geçerli bir tarih olmalıdır."
        return _result
    _result = True
    return _result
    errorMsg.value = "İzlanda VKN'sinin ilk 6 rakamı geçerli bir tarih olmalıdır."
    return _result

def ValidateIndiaPAN(pan, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanPAN = ""
    cleanPAN = Replace(Trim(pan), " ", "")
    cleanPAN = UCase(cleanPAN)
    if Len(cleanPAN) != 10:
        errorMsg.value = "Hindistan PAN'ı 10 karakter olmalıdır."
        return _result
    if not vba_like(Mid(cleanPAN, 1, 3), "[A-Z][A-Z][A-Z]"):
        errorMsg.value = "Hindistan PAN'ının ilk 3 karakteri harf olmalıdır (A-Z)."
        return _result
    if Mid(cleanPAN, 4, 1) == "P" or Mid(cleanPAN, 4, 1) == "F" or Mid(cleanPAN, 4, 1) == "C" or Mid(cleanPAN, 4, 1) == "H" or Mid(cleanPAN, 4, 1) == "A" or Mid(cleanPAN, 4, 1) == "T" or Mid(cleanPAN, 4, 1) == "B" or Mid(cleanPAN, 4, 1) == "L" or Mid(cleanPAN, 4, 1) == "J" or Mid(cleanPAN, 4, 1) == "G":
        pass
    else:
        errorMsg.value = "Hindistan PAN'ının 4. karakteri geçerli bir durum kodu olmalıdır (P, F, C, H, A, T, B, L, J, G)."
        return _result
    if not vba_like(Mid(cleanPAN, 5, 1), "[A-Z]"):
        errorMsg.value = "Hindistan PAN'ının 5. karakteri harf olmalıdır (A-Z)."
        return _result
    if not vba_like(Mid(cleanPAN, 6, 4), "####"):
        errorMsg.value = "Hindistan PAN'ının 6-9 arasındaki karakterleri rakamlardan oluşmalıdır (0001-9999)."
        return _result
    if not vba_like(Mid(cleanPAN, 10, 1), "[A-Z]"):
        errorMsg.value = "Hindistan PAN'ının 10. karakteri harf olmalıdır (A-Z)."
        return _result
    _result = True
    return _result

def ValidateIndonesiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Trim(tin)
    if Len(tin) != 15 and Len(tin) != 16:
        errorMsg.value = "Endonezya VKN'si 15 veya 16 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Endonezya VKN'si sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateIraqTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Irak TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9 and Len(cleanTIN) != 10:
        errorMsg.value = "Irak TIN'i 9 veya 10 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateIrelandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinClean = ""
    tinLength = 0
    digitsPart = ""
    lettersPart = ""
    tinClean = Trim(tin)
    tinLength = Len(tinClean)
    if tinLength < 8 or tinLength > 9:
        errorMsg.value = "İrlanda VKN'si 8 veya 9 karakter olmalıdır (7 rakam + 1 veya 2 harf)."
        return _result
    digitsPart = Left(tinClean, 7)
    lettersPart = Mid(tinClean, 8)
    if not IsAllNumeric(digitsPart):
        errorMsg.value = "İrlanda VKN'sinin ilk 7 karakteri rakamlardan oluşmalıdır."
        return _result
    if not IsAllLetters(lettersPart):
        errorMsg.value = "İrlanda VKN'sinin son 1 veya 2 karakteri harflerden oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateIsleOfManTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    prefix = ""
    digitsPart = ""
    suffix = ""
    isValidFormat = False
    tinType = ""
    cleanTIN = ""
    suffixLetter = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) == 7 or Len(cleanTIN) == 9:
        prefix = Left(cleanTIN, 1)
        if prefix == "H" or prefix == "C" or prefix == "X":
            digitsPart = Mid(cleanTIN, 2, 6)
            if Len(digitsPart) == 6 and IsNumeric(digitsPart):
                if Len(cleanTIN) == 9:
                    suffix = Mid(cleanTIN, 8, 2)
                    if not (Len(suffix) == 2 and IsNumeric(suffix)):
                        errorMsg.value = "İsteğe bağlı ek olan son 2 hane sayı olmalıdır."
                        return _result
                _result = True
                return _result
            else:
                errorMsg.value = "6 haneli sayısal bir bölüm bulunmalıdır."
                return _result
        else:
            errorMsg.value = "Geçersiz ön ek. Ön ek H, C veya X olmalıdır."
            return _result
    elif Len(cleanTIN) == 9:
        if vba_like(Mid(cleanTIN, 1, 2), "[A-Z][A-Z]"):
            digitsPart = Mid(cleanTIN, 3, 6)
            if IsNumeric(digitsPart):
                suffixLetter = Mid(cleanTIN, 9, 1)
                if suffixLetter == "A" or suffixLetter == "B" or suffixLetter == "C" or suffixLetter == "D":
                    _result = True
                    return _result
                else:
                    errorMsg.value = "Son karakter A, B, C veya D olmalıdır."
                    return _result
            else:
                errorMsg.value = "Orta 6 karakter sayısal olmalıdır."
                return _result
        else:
            errorMsg.value = "İlk iki karakter harf olmalıdır."
            return _result
    else:
        errorMsg.value = "Geçersiz TIN formatı."
        return _result
    errorMsg.value = "Geçersiz Man Adaları TIN'i."
    _result = False
    return _result

def ValidateIsraelTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Trim(tin)
    if Len(tin) != 9:
        errorMsg.value = "İsrail VKN'si tam olarak 9 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "İsrail VKN'si sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateItalyVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    j = 0
    if Len(vkn) == 16:
        for j in vba_range(1, 16, 1):
            if not (vba_like(Mid(vkn, j, 1), "[A-Za-z0-9]")):
                errorMsg.value = "İtalya Codice Fiscale sadece harf ve rakamlardan oluşmalıdır."
                return _result
        _result = True
        return _result
    elif Len(vkn) == 11:
        for j in vba_range(1, 11, 1):
            if not (vba_like(Mid(vkn, j, 1), "[0-9]")):
                errorMsg.value = "İtalya Partita IVA sadece rakamlardan oluşmalıdır."
                return _result
        _result = True
        return _result
    else:
        errorMsg.value = "İtalya VKN'si ya 16 ya da 11 karakter uzunluğunda olmalıdır."
    return _result

def ValidateJamaicaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tin = Trim(tin)
    if Len(tin) != 9:
        errorMsg.value = "Jamaika VKN'si tam olarak 9 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Jamaika VKN'si sadece rakamlardan oluşmalıdır."
        return _result
    if Left(tin, 1) != "0" and Left(tin, 1) != "1":
        errorMsg.value = "Jamaika VKN'sinin ilk rakamı 0 veya 1 olmalıdır."
        return _result
    _result = True
    return _result

def ValidateJapanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = ""
    i = 0
    for i in vba_range(1, Len(tin), 1):
        if IsNumeric(Mid(tin, i, 1)):
            cleanTIN = cleanTIN + Mid(tin, i, 1)
    if Len(cleanTIN) == 12 or Len(cleanTIN) == 13:
        _result = True
    else:
        errorMsg.value = "Japonya TIN'i 12 haneli (Bireysel Numara) veya 13 haneli (Kurumsal Numara) olmalıdır (sadece rakamlar)."
    return _result

def ValidateJerseyTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    tinClean = ""
    i = 0
    hyphenPositions = None
    expectedHyphenPositions = None
    hyphenCount = 0
    tinLength = 0
    tin = Trim(tin)
    tinClean = Replace(Replace(tin, "-", ""), " ", "")
    if Len(tinClean) != 10:
        errorMsg.value = "Jersey VKN'si 10 haneli bir sayı olmalıdır."
        return _result
    if not IsAllNumeric(tinClean):
        errorMsg.value = "Jersey VKN'si sadece rakamlardan oluşmalıdır."
        return _result
    expectedHyphenPositions = Array(4, 8)
    hyphenCount = 0
    for i in vba_range(1, Len(tin), 1):
        if Mid(tin, i, 1) == "-":
            hyphenCount = hyphenCount + 1
            if hyphenCount <= UBound(expectedHyphenPositions) + 1:
                if i != expectedHyphenPositions(hyphenCount - 1):
                    errorMsg.value = "Jersey VKN'sindeki tireler doğru pozisyonda olmalıdır: XXX-XXX-XXXX"
                    return _result
            else:
                errorMsg.value = "Jersey VKN'sindeki tire sayısı fazla."
                return _result
    _result = True
    return _result

def ValidateKazakhstanIIN(iin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    iin = Replace(Trim(iin), " ", "")
    if not IsAllNumeric(iin):
        errorMsg.value = "Kazakistan IIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(iin) != 12:
        errorMsg.value = "Kazakistan IIN'i 12 haneli olmalıdır."
        return _result
    weights = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)
    weightedSum = 0
    for i in vba_range(1, 11, 1):
        digit = CInt(Mid(iin, i, 1))
        weightedSum = weightedSum + digit * weights(i - 1)
    calculatedCheckDigit = weightedSum % 11
    if calculatedCheckDigit == 10:
        weights = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2)
        weightedSum = 0
        for i in vba_range(1, 11, 1):
            digit = CInt(Mid(iin, i, 1))
            weightedSum = weightedSum + digit * weights(i - 1)
        calculatedCheckDigit = weightedSum % 11
        if calculatedCheckDigit == 10:
            calculatedCheckDigit = 0
    checkDigit = CInt(Right(iin, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Kazakistan IIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateKazakhstanBIN(bin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    bin = Replace(Trim(bin), " ", "")
    if not IsAllNumeric(bin):
        errorMsg.value = "Kazakistan BIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(bin) != 12:
        errorMsg.value = "Kazakistan BIN'i 12 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateKenyaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = UCase(Replace(Replace(Trim(tin), " ", ""), "-", ""))
    if Len(cleanTIN) != 11:
        errorMsg.value = "Kenya TIN'i 11 karakter olmalıdır."
        return _result
    if not (vba_like(Left(cleanTIN, 1), "[A-Z]") and vba_like(Right(cleanTIN, 1), "[A-Z]") and vba_like(Mid(cleanTIN, 2, 9), "#########")):
        errorMsg.value = "Kenya TIN'i geçersiz formatta. İlk ve son karakterler harf, diğerleri rakam olmalıdır."
        return _result
    if Left(cleanTIN, 1) == "P" or Left(cleanTIN, 1) == "A":
        _result = True
        return _result
    else:
        errorMsg.value = "Kenya TIN'i 'P'(tüzel kişi) veya 'A' (gerçek kişi) ile başlamalıdır."
    return _result

def ValidateKosovoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Kosova TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Kosova TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateKuwaitTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanedTIN = ""
    tinLength = 0
    isIndividual = False
    isEntity = False
    i = 0
    centuryCode = 0
    yearPart = ""
    monthPart = ""
    dayPart = ""
    fullYear = 0
    fullDate = None
    serialNumber = ""
    placeCode = ""
    cleanedTIN = ""
    for i in vba_range(1, Len(tin), 1):
        if vba_like(Mid(tin, i, 1), "#"):
            cleanedTIN = cleanedTIN + Mid(tin, i, 1)
    tinLength = Len(cleanedTIN)
    if not IsNumeric(cleanedTIN):
        errorMsg.value = "Kuveyt TIN'i sadece rakamlardan oluşmalıdır. Harf, boşluk veya özel karakter içeremez."
        return _result
    if tinLength == 12:
        isIndividual = True
    elif tinLength == 6:
        isEntity = True
    else:
        errorMsg.value = "Kuveyt TIN'i bireyler için 12 haneli, tüzel kişiler için 6 haneli olmalıdır."
        return _result
    if isIndividual:
        centuryCode = CInt(Mid(cleanedTIN, 1, 1))
        if centuryCode != 2 and centuryCode != 3:
            errorMsg.value = "Geçersiz yüzyıl kodu. Bireylerin TIN'leri 2 veya 3 ile başlamalıdır."
            return _result
        yearPart = Mid(cleanedTIN, 2, 2)
        monthPart = Mid(cleanedTIN, 4, 2)
        dayPart = Mid(cleanedTIN, 6, 2)
        if not (IsNumeric(yearPart) and IsNumeric(monthPart) and IsNumeric(dayPart)):
            errorMsg.value = "Doğum tarihi bölümleri sayısal olmalıdır."
            return _result
        if centuryCode == 2:
            fullYear = 1900 + CInt(yearPart)
        elif centuryCode == 3:
            fullYear = 2000 + CInt(yearPart)
        fullDate = DateSerial(fullYear, CInt(monthPart), CInt(dayPart))
        placeCode = Mid(cleanedTIN, 8, 1)
        if not IsNumeric(placeCode):
            errorMsg.value = "Doğum yeri kodu sayısal olmalıdır."
            return _result
        serialNumber = Mid(cleanedTIN, 9, 4)
        if not IsNumeric(serialNumber):
            errorMsg.value = "Seri numarası sayısal olmalıdır."
            return _result
        _result = True
        return _result
    elif isEntity:
        _result = True
        return _result
    errorMsg.value = "Geçersiz doğum tarihi."
    _result = False
    return _result

def ValidateLatviaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    centuryDigit = ""
    dayPart = 0
    monthPart = 0
    yearPart = 0
    fullYear = 0
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Letonya TIN'i sadece rakamlardan oluşmalıdır."
        _result = False
        return _result
    if Len(cleanTIN) != 11:
        errorMsg.value = "Letonya TIN'i 11 haneli olmalıdır."
        _result = False
        return _result
    if Left(cleanTIN, 2) == "32":
        _result = True
        return _result
    dayPart = CInt(Mid(cleanTIN, 1, 2))
    monthPart = CInt(Mid(cleanTIN, 3, 2))
    yearPart = CInt(Mid(cleanTIN, 5, 2))
    centuryDigit = Mid(cleanTIN, 7, 1)
    if centuryDigit == "0":
        fullYear = 1800 + yearPart
    elif centuryDigit == "1":
        fullYear = 1900 + yearPart
    elif centuryDigit == "2":
        fullYear = 2000 + yearPart
    else:
        errorMsg.value = "Letonya TIN'inin 7. hanesi geçersiz yüzyıl kodudur (0, 1 veya 2 olmalıdır)."
        _result = False
        return _result
    if monthPart < 1 or monthPart > 12:
        errorMsg.value = "Letonya TIN'inde geçersiz ay bilgisi var (01-12 olmalıdır)."
        _result = False
        return _result
    if not IsValidDay(dayPart, monthPart, fullYear):
        errorMsg.value = "Letonya TIN'inde geçersiz gün bilgisi var."
        _result = False
        return _result
    _result = True
    return _result
    return _result

def ValidateLebanonTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) > 16:
        errorMsg.value = "Lübnan TIN'i en fazla 16 karakter olabilir."
        return _result
    _result = True
    return _result

def ValidateLesothoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Trim(tin)
    if Len(cleanTIN) != 9 and Len(cleanTIN) != 10:
        errorMsg.value = "Lesotho TIN'i 9 veya 10 karakter olmalıdır."
        return _result
    _result = True
    return _result

def ValidateLibyaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) > 16:
        errorMsg.value = "Libya TIN'i en fazla 16 karakter olabilir."
        return _result
    _result = True
    return _result

def ValidateLiechtensteinTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = ""
    i = 0
    for i in vba_range(1, Len(tin), 1):
        if IsNumeric(Mid(tin, i, 1)):
            cleanTIN = cleanTIN + Mid(tin, i, 1)
    if Len(cleanTIN) > 0 and Len(cleanTIN) <= 12:
        _result = True
    else:
        errorMsg.value = "Liechtenstein TIN'i 1 ile 12 arasında rakam içermelidir."
    return _result

def ValidateLithuaniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights1 = VbaArray(1, 10, 0)
    weights2 = VbaArray(1, 10, 0)
    sum = 0
    remainder = 0
    firstDigit = 0
    yearPart = 0
    monthPart = 0
    dayPart = 0
    fullYear = 0
    birthdate = None
    errorMsg.value = ""
    _result = False
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) != 11:
        errorMsg.value = "Litvanya TIN'i 11 haneli olmalıdır."
        return _result
    if not IsAllDigits(cleanTIN):
        errorMsg.value = "Litvanya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    firstDigit = CInt(Mid(cleanTIN, 1, 1))
    if firstDigit < 1 or firstDigit > 6:
        errorMsg.value = "Litvanya TIN'inin ilk hanesi 1 ile 6 arasında olmalıdır."
        return _result
    yearPart = CInt(Mid(cleanTIN, 2, 2))
    monthPart = CInt(Mid(cleanTIN, 4, 2))
    dayPart = CInt(Mid(cleanTIN, 6, 2))
    if firstDigit == 1 or firstDigit == 2:
        fullYear = 1800 + yearPart
    elif firstDigit == 3 or firstDigit == 4:
        fullYear = 1900 + yearPart
    elif firstDigit == 5 or firstDigit == 6:
        fullYear = 2000 + yearPart
    else:
        errorMsg.value = "Litvanya TIN'inin ilk hanesi geçersiz."
        return _result
    if monthPart < 1 or monthPart > 12:
        errorMsg.value = "Litvanya TIN'inde geçersiz ay bilgisi var (01-12 olmalıdır)."
        return _result
    birthdate = DateSerial(fullYear, monthPart, dayPart)
    weights1[1] = 1
    weights1[2] = 2
    weights1[3] = 3
    weights1[4] = 4
    weights1[5] = 5
    weights1[6] = 6
    weights1[7] = 7
    weights1[8] = 8
    weights1[9] = 9
    weights1[10] = 1
    weights2[1] = 3
    weights2[2] = 4
    weights2[3] = 5
    weights2[4] = 6
    weights2[5] = 7
    weights2[6] = 8
    weights2[7] = 9
    weights2[8] = 1
    weights2[9] = 2
    weights2[10] = 3
    sum = 0
    for i in vba_range(1, 10, 1):
        sum = sum + CInt(Mid(cleanTIN, i, 1)) * weights1(i)
    remainder = sum % 11
    if remainder != 10:
        calculatedCheckDigit = remainder
    else:
        sum = 0
        for i in vba_range(1, 10, 1):
            sum = sum + CInt(Mid(cleanTIN, i, 1)) * weights2(i)
        remainder = sum % 11
        if remainder != 10:
            calculatedCheckDigit = remainder
        else:
            calculatedCheckDigit = 0
    checkDigit = CInt(Mid(cleanTIN, 11, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Litvanya TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result
    errorMsg.value = "Litvanya TIN'inde geçersiz doğum tarihi var."
    _result = False
    return _result

def ValidateMacaoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Macao TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 8:
        errorMsg.value = "Macao TIN'i 8 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateMadagascarTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) > 16:
        errorMsg.value = "Madagaskar TIN'i en fazla 16 karakter olabilir."
        return _result
    _result = True
    return _result

def ValidateMalaysiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    isIndividual = False
    isNonIndividual = False
    isNRIC = False
    cleanTIN = Trim(tin)
    if UCase(Left(cleanTIN, 2)) == "IG":
        afterIG = ""
        afterIG = Mid(cleanTIN, 3)
        if Len(afterIG) == 9 and IsAllNumeric(afterIG):
            _result = True
            return _result
        else:
            errorMsg.value = "Malezya Bireysel TIN formatı hatalı. 'IG' sonrası tam 9 rakam olmalı."
            return _result
    validNonIndCodes = None
    validNonIndCodes = Array("C", "D", "E", "F", "G", "J", "FA", "PT", "TA", "TC", "TN", "TR", "TP", "LE")
    codeFound = False
    codeLength = 0
    nonIndCode = ""
    codeFound = False
    if Len(cleanTIN) > 2:
        nonIndCode = UCase(Left(cleanTIN, 2))
        if not IsError(Application.Match(nonIndCode, validNonIndCodes, 0)):
            codeFound = True
            codeLength = 2
    if not codeFound and Len(cleanTIN) > 1:
        nonIndCode = UCase(Left(cleanTIN, 1))
        if not IsError(Application.Match(nonIndCode, validNonIndCodes, 0)):
            codeFound = True
            codeLength = 1
    if codeFound:
        if Right(cleanTIN, 1) != "0":
            errorMsg.value = "Malezya Non-Individual TIN son karakter '0' ile bitmelidir."
            return _result
        middlePart = ""
        middlePart = Mid(cleanTIN, codeLength + 1, Len(cleanTIN) - codeLength - 1)
        if middlePart == "" or not IsAllNumeric(middlePart):
            errorMsg.value = "Malezya Non-Individual TIN, kod sonrası rakamlar ve son karakter '0' formatında olmalıdır."
            return _result
        _result = True
        return _result
    if Len(cleanTIN) == 12 and IsAllNumeric(cleanTIN):
        _result = True
        return _result
    errorMsg.value = "Geçersiz Malezya TIN. Bireysel (IG...), Non-Individual (kod+...+0) veya 12 haneli NRIC giriniz."
    return _result

def ValidateMaldivesTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    bpPart = ""
    revPart = ""
    seqPart = ""
    i = 0
    cleanTIN = Trim(tin)
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) < 12 or Len(cleanTIN) > 15:
        errorMsg.value = "Maldivler TIN uzunluğu geçerli değil (12-15 karakter arası olmalı)."
        return _result
    bpPart = Left(cleanTIN, 7)
    if not IsAllNumeric(bpPart):
        errorMsg.value = "Maldivler TIN'inin ilk 7 karakteri rakamlardan oluşmalıdır."
        return _result
    seqPart = Right(cleanTIN, 3)
    if not IsAllNumeric(seqPart):
        errorMsg.value = "Maldivler TIN'inin son 3 karakteri rakamlardan oluşmalıdır."
        return _result
    revLength = 0
    revLength = Len(cleanTIN) - 10
    revPart = Mid(cleanTIN, 8, revLength)
    if revLength < 2 or revLength > 5:
        errorMsg.value = "Maldivler TIN'inin Revenue Code'u 2 ile 5 karakter uzunluğunda olmalıdır."
        return _result
    for i in vba_range(1, Len(revPart), 1):
        if not (vba_like(Mid(revPart, i, 1), "[A-Z]")):
            errorMsg.value = "Maldivler TIN'inin Revenue Code bölümünde sadece harf bulunmalıdır."
            return _result
    _result = True
    return _result

def ValidateMaltaFormat1_Final(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = Trim(tin)
    if Len(cleanTIN) != 8:
        errorMsg.value = "Malta TIN format 1 must be 8 characters."
        return _result
    if not IsAllNumeric(Left(cleanTIN, 7)) or not vba_like(Right(cleanTIN, 1), "[A-Z]"):
        errorMsg.value = "Malta TIN format 1 must contain 7 digits followed by one letter."
        return _result
    _result = True
    return _result


def ValidateMaltaFormat2_Final(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = Trim(tin)
    validPrefixes = ("11", "22", "33", "44", "55", "66", "77", "88")
    if Len(cleanTIN) != 9:
        errorMsg.value = "Malta TIN format 2 must be 9 digits."
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Malta TIN format 2 must contain digits only."
        return _result
    if Left(cleanTIN, 2) not in validPrefixes:
        errorMsg.value = "Malta TIN format 2 must start with 11, 22, 33, 44, 55, 66, 77 or 88."
        return _result
    _result = True
    return _result


def ValidateMaltaTIN_Final(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Trim(tin)
    if Len(cleanTIN) == 8:
        _result = ValidateMaltaFormat1_Final(cleanTIN, errorMsg)
    elif Len(cleanTIN) == 9:
        _result = ValidateMaltaFormat2_Final(cleanTIN, errorMsg)
    else:
        errorMsg.value = "Malta TIN uzunluğu geçersiz. Malta 1 için 8 karakter, Malta 2 için 9 karakter olmalı."
    return _result

def ValidateMarshallIslandsTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Trim(tin)
    if Len(cleanTIN) == 8:
        if Mid(cleanTIN, 6, 1) == "-":
            firstFive = ""
            firstFive = Left(cleanTIN, 5)
            if IsAllNumeric(firstFive) and Right(cleanTIN, 2) == "04":
                _result = True
                return _result
            else:
                pass
    if Len(cleanTIN) == 9:
        if Left(cleanTIN, 2) == "04" and Mid(cleanTIN, 3, 1) == "-":
            lastSix = ""
            lastSix = Right(cleanTIN, 6)
            if IsAllNumeric(lastSix):
                _result = True
                return _result
    errorMsg.value = "Marshall Adaları TIN formatı geçersiz. EIN: #####-04 veya Çalışan No: 04-###### olmalı."
    return _result

def ValidateMauritiusTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Trim(tin)
    if Len(cleanTIN) != 8:
        errorMsg.value = "Mauritius TAN 8 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Mauritius TAN sadece rakamlardan oluşmalıdır."
        return _result
    firstDigit = ""
    firstDigit = Left(cleanTIN, 1)
    if firstDigit == "1" or firstDigit == "5" or firstDigit == "7" or firstDigit == "8":
        _result = True
    elif firstDigit == "2" or firstDigit == "3":
        _result = True
    else:
        errorMsg.value = "Mauritius TAN'in ilk rakamı bireyler için 1,5,7,8; tüzel kişiler için 2 veya 3 olmalıdır."
    return _result

def ValidateMexicoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    firstPart = ""
    datePart = ""
    homoclave = ""
    isIndividual = False
    year = 0
    month = 0
    day = 0
    dateValue = None
    if Len(tin) == 13:
        isIndividual = True
    elif Len(tin) == 12:
        isIndividual = False
    else:
        errorMsg.value = "Meksika TIN'i bireyler için 13 karakter, tüzel kişiler için 12 karakter olmalıdır."
        return _result
    for i in vba_range(1, Len(tin), 1):
        if not vba_like(Mid(tin, i, 1), "[A-Z0-9]"):
            errorMsg.value = "Meksika TIN'i yalnızca harfler ve rakamlardan oluşmalıdır. Boşluk veya özel karakter içeremez."
            return _result
    if isIndividual:
        firstPart = Left(tin, 4)
        datePart = Mid(tin, 5, 6)
        homoclave = Right(tin, 3)
        if not vba_like(firstPart, "[A-Z][A-Z][A-Z][A-Z]"):
            errorMsg.value = "Bireyler için Meksika TIN'inin ilk 4 karakteri harf olmalıdır."
            return _result
    else:
        firstPart = Left(tin, 3)
        datePart = Mid(tin, 4, 6)
        homoclave = Right(tin, 3)
        if not vba_like(firstPart, "[A-Z][A-Z][A-Z]"):
            errorMsg.value = "Tüzel kişiler için Meksika TIN'inin ilk 3 karakteri harf olmalıdır."
            return _result
    if not vba_like(datePart, "######"):
        errorMsg.value = "Meksika TIN'inin tarih kısmı (6 hane) sadece rakamlardan oluşmalıdır."
        return _result
    year = CInt(Left(datePart, 2))
    month = CInt(Mid(datePart, 3, 2))
    day = CInt(Right(datePart, 2))
    if year >= 0 and year <= 99:
        if year >= 0 and year <= 21:
            year = 2000 + year
        else:
            year = 1900 + year
    else:
        errorMsg.value = "Meksika TIN'inde yıl kısmı geçersiz."
        return _result
    dateValue = DateSerial(year, month, day)
    if not vba_like(homoclave, "[A-Z0-9][A-Z0-9][A-Z0-9]"):
        errorMsg.value = "Meksika TIN'inin son 3 karakteri (homoclave) harf veya rakam olmalıdır."
        return _result
    _result = True
    return _result
    errorMsg.value = "Meksika TIN'inde tarih kısmı geçersiz."
    return _result
    return _result

def ValidateMoldovaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Moldova TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 7 and Len(cleanTIN) != 8 and Len(cleanTIN) != 13:
        errorMsg.value = "Moldova TIN'i 7, 8 veya 13 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateMonacoSIRET(siret, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanSIRET = ""
    cleanSIRET = Replace(Trim(siret), " ", "")
    if not IsAllNumeric(cleanSIRET):
        errorMsg.value = "Monako SIRET numarası sadece rakamlardan oluşmalıdır. Sadece tüzel kişiler için."
        return _result
    if Len(cleanSIRET) != 14:
        errorMsg.value = "Monako SIRET numarası 14 haneli olmalıdır. Sadece tüzel kişiler için."
        return _result
    _result = True
    return _result

def ValidateMongoliaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Moğolistan TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 7:
        errorMsg.value = "Moğolistan TIN'i 7 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateMontenegroTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Karadağ TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 8 and Len(cleanTIN) != 13:
        errorMsg.value = "Karadağ TIN'i 8 veya 13 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateMoroccoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Fas TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 15:
        errorMsg.value = "Fas TIN'i 15 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateMozambiqueTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Mozambik TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Mozambik TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateNamibiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Namibya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 8:
        errorMsg.value = "Namibya TIN'i 8 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateNauruTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) != 9:
        errorMsg.value = "Nauru TIN'i 9 haneli olmalıdır."
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Nauru TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateNepalTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Nepal TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Nepal TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateNetherlandsTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Hollanda TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Hollanda TIN'i 9 haneli olmalıdır."
        return _result
    weights = Array(9, 8, 7, 6, 5, 4, 3, 2)
    weightedSum = 0
    for i in vba_range(1, 8, 1):
        digit = CInt(Mid(cleanTIN, i, 1))
        weightedSum = weightedSum + digit * weights(i - 1)
    calculatedCheckDigit = weightedSum % 11
    checkDigit = CInt(Right(cleanTIN, 1))
    if calculatedCheckDigit == 10:
        errorMsg.value = "Geçersiz Hollanda TIN. Modulo sonucu 10 olamaz"
        return _result
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Hollanda TIN. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateNewZealandIRD(ird, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanIRD = ""
    cleanIRD = Replace(Replace(Trim(ird), " ", ""), "-", "")
    if Len(cleanIRD) != 8 and Len(cleanIRD) != 9:
        errorMsg.value = "Yeni Zelanda IRD numarası 8 veya 9 haneli olmalıdır."
        return _result
    if not IsAllNumeric(cleanIRD):
        errorMsg.value = "Yeni Zelanda IRD numarası sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateNigeriaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if Len(cleanTIN) == 12:
        if Right(cleanTIN, 4) == "0001" and IsAllNumeric(Left(cleanTIN, 8)):
            _result = True
            return _result
        else:
            errorMsg.value = "Nijerya FIRS TIN'i 8 rakam ve ardından '-0001' olmalıdır. Örnek: 12345678-0001"
            return _result
    elif Len(cleanTIN) == 10:
        if IsAllNumeric(cleanTIN):
            _result = True
            return _result
        else:
            errorMsg.value = "Nijerya JTB TIN'i 10 rakamdan oluşmalıdır."
            return _result
    else:
        errorMsg.value = "Nijerya TIN'i geçersiz. 10 haneli JTB TIN veya 8 hane + '-0001' olan FIRS TIN formatında olmalıdır."
    return _result

def ValidateNorthMacedoniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Left(cleanTIN, 2) == "MK":
        numericPart = Right(cleanTIN, 13)
        if Len(cleanTIN) != 15:
            errorMsg.value = "Kuzey Makedonya TIN'i 'MK' ön ekiyle başlarsa 15 karakter uzunluğunda olmalıdır."
            return _result
    else:
        numericPart = cleanTIN
        if Len(cleanTIN) != 13:
            errorMsg.value = "Kuzey Makedonya TIN'i 13 haneli olmalıdır veya 'MK' ile başlamalı ve ardından 13 rakam gelmelidir."
            return _result
    if not IsAllNumeric(numericPart):
        errorMsg.value = "Kuzey Makedonya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateNorwayTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    birthdate = None
    dayPart = 0
    monthPart = 0
    yearPart = 0
    individualNumber = 0
    fullYear = 0
    k1 = 0
    k2 = 0
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) == 11:
        if IsAllNumeric(cleanTIN):
            dayPart = CInt(Mid(cleanTIN, 1, 2))
            monthPart = CInt(Mid(cleanTIN, 3, 2))
            yearPart = CInt(Mid(cleanTIN, 5, 2))
            individualNumber = CInt(Mid(cleanTIN, 7, 3))
            if individualNumber >= 0 and individualNumber <= 499:
                fullYear = 1900 + yearPart
            elif individualNumber >= 500 and individualNumber <= 749 and yearPart >= 54 and yearPart <= 99:
                fullYear = 1800 + yearPart
            elif individualNumber >= 500 and individualNumber <= 999 and yearPart >= 0 and yearPart <= 39:
                fullYear = 2000 + yearPart
            elif individualNumber >= 900 and individualNumber <= 999 and yearPart >= 40 and yearPart <= 99:
                fullYear = 1900 + yearPart
            else:
                errorMsg.value = "Invalid individual number in National Identity Number."
                return _result
            birthdate = DateSerial(fullYear, monthPart, dayPart)
            digits = VbaArray(1, 11, 0)
            for i in vba_range(1, 11, 1):
                digits[i] = CInt(Mid(cleanTIN, i, 1))
            weights = Array(0, 3, 7, 6, 1, 8, 9, 4, 5, 2)
            weightedSum = 0
            for i in vba_range(1, 9, 1):
                weightedSum = weightedSum + digits(i) * weights(i)
            k1 = 11 - (weightedSum % 11)
            if k1 == 11:
                k1 = 0
            if k1 == 10:
                errorMsg.value = "Invalid control digit K1 in National Identity Number."
                return _result
            if k1 != digits(10):
                errorMsg.value = "Control digit K1 does not match in National Identity Number."
                return _result
            weights = Array(0, 5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
            weightedSum = 0
            for i in vba_range(1, 10, 1):
                weightedSum = weightedSum + digits(i) * weights(i)
            k2 = 11 - (weightedSum % 11)
            if k2 == 11:
                k2 = 0
            if k2 == 10:
                errorMsg.value = "Invalid control digit K2 in National Identity Number."
                return _result
            if k2 != digits(11):
                errorMsg.value = "Control digit K2 does not match in National Identity Number."
                return _result
            _result = True
            return _result
            errorMsg.value = "Invalid date in National Identity Number."
            return _result
        else:
            errorMsg.value = "Norveç Ulusal Kimlik Numarası 11 haneli ve sadece rakamlardan oluşmalıdır."
            return _result
    elif Len(cleanTIN) == 9:
        if IsAllNumeric(cleanTIN):
            if Left(cleanTIN, 1) == "8" or Left(cleanTIN, 1) == "9":
                weights = Array(3, 2, 7, 6, 5, 4, 3, 2)
                weightedSum = 0
                for i in vba_range(1, 8, 1):
                    digit = CInt(Mid(cleanTIN, i, 1))
                    weightedSum = weightedSum + digit * weights(i - 1)
                calculatedCheckDigit = 11 - (weightedSum % 11)
                if calculatedCheckDigit == 11:
                    calculatedCheckDigit = 0
                if calculatedCheckDigit == 10:
                    errorMsg.value = "Geçersiz kontrol basamağı Norveç Organizasyon Numarasında."
                    return _result
                checkDigit = CInt(Right(cleanTIN, 1))
                if calculatedCheckDigit != checkDigit:
                    errorMsg.value = "Geçersiz Norveç Organizasyon Numarası. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
                    return _result
                _result = True
                return _result
            else:
                errorMsg.value = "Norveç Organizasyon numarası 8 veya 9 ile başlamalıdır."
                return _result
        else:
            errorMsg.value = "Norveç Organizasyon Numarası 9 haneli ve sadece rakamlardan oluşmalıdır."
            return _result
    elif UCase(Right(cleanTIN, 3)) == "MVA" and Len(cleanTIN) > 3:
        orgNumber = ""
        orgNumber = Left(cleanTIN, Len(cleanTIN) - 3)
        if Len(orgNumber) == 9 and IsAllNumeric(orgNumber):
            if Left(orgNumber, 1) == "8" or Left(orgNumber, 1) == "9":
                weights = Array(3, 2, 7, 6, 5, 4, 3, 2)
                weightedSum = 0
                for i in vba_range(1, 8, 1):
                    digit = CInt(Mid(orgNumber, i, 1))
                    weightedSum = weightedSum + digit * weights(i - 1)
                calculatedCheckDigit = 11 - (weightedSum % 11)
                if calculatedCheckDigit == 11:
                    calculatedCheckDigit = 0
                if calculatedCheckDigit == 10:
                    errorMsg.value = "Geçersiz kontrol basamağı Norveç Organizasyon Numarasında."
                    return _result
                checkDigit = CInt(Right(orgNumber, 1))
                if calculatedCheckDigit != checkDigit:
                    errorMsg.value = "Geçersiz Norveç Organizasyon Numarası KDV Kayıtlı Tüzel Kişide. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
                    return _result
                _result = True
                return _result
            else:
                errorMsg.value = "KDV Kayıtlı Tüzel Kişiler için Norveç Organizasyon Numarası 8 veya 9 ile başlamalıdır."
                return _result
        else:
            errorMsg.value = "KDV Kayıtlı Tüzel Kişiler için Norveç TIN'i 9 haneli bir sayı ve 'MVA' sonekinden oluşmalıdır."
            return _result
    else:
        errorMsg.value = "Norveç TIN'i geçersiz."
        return _result
    return _result

def ValidateOmanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) > 7:
        errorMsg.value = "Oman TIN'i en fazla 7 haneli olmalıdır."
        return _result
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Oman TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) == 0:
        errorMsg.value = "Oman TIN'i boş olamaz."
        return _result
    _result = True
    return _result

def ValidatePakistanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    if Len(tin) == 7 or Len(tin) == 13:
        for i in vba_range(1, Len(tin), 1):
            if not vba_like(Mid(tin, i, 1), "#"):
                errorMsg.value = "Pakistan TIN'i sadece rakamlardan oluşmalıdır. Boşluk, harf veya özel karakter içeremez."
                return _result
        _result = True
    else:
        errorMsg.value = "Pakistan TIN'i 7 veya 13 haneli bir sayı olmalıdır."
    return _result

def ValidatePeruTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    if Len(tin) != 11:
        errorMsg.value = "Peru TIN'i tam olarak 11 rakamdan oluşmalıdır."
        return _result
    for i in vba_range(1, Len(tin), 1):
        if not vba_like(Mid(tin, i, 1), "#"):
            errorMsg.value = "Peru TIN'i sadece rakamlardan oluşmalıdır. Boşluk, harf veya özel karakter içeremez."
            return _result
    _result = True
    return _result

def ValidatePhilippinesTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    firstDigit = ""
    isIndividual = False
    isPartnership = False
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Filipinler TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9 and Len(cleanTIN) != 12:
        errorMsg.value = "Filipinler TIN'i 9 veya 12 haneli olmalıdır."
        return _result
    firstDigit = Left(cleanTIN, 1)
    if vba_like(firstDigit, "[1-9]"):
        _result = True
        return _result
    elif firstDigit == "0":
        _result = True
        return _result
    else:
        errorMsg.value = "Filipinler TIN'i bireyler için 1-9 ile, kuruluşlar için 0 ile başlamalıdır."
        return _result
    return _result

def ValidatePolandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    tinLength = 0
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    monthPart = 0
    century = 0
    yearPart = 0
    monthTable = None
    fullYear = 0
    dayPart = 0
    birthdate = None
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Polonya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    tinLength = Len(cleanTIN)
    if tinLength != 10 and tinLength != 11:
        errorMsg.value = "Polonya TIN'i 10 veya 11 haneli olmalıdır."
        return _result
    if tinLength == 10:
        weights = Array(6, 5, 7, 2, 3, 4, 5, 6, 7)
        weightedSum = 0
        for i in vba_range(1, 9, 1):
            digit = CInt(Mid(cleanTIN, i, 1))
            weightedSum = weightedSum + digit * weights(i - 1)
        calculatedCheckDigit = weightedSum % 11
        if calculatedCheckDigit == 10:
            calculatedCheckDigit = 0
        checkDigit = CInt(Right(cleanTIN, 1))
        if calculatedCheckDigit != checkDigit:
            errorMsg.value = "Geçersiz Polonya TIN (Yapı 1). Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
            return _result
        _result = True
    elif tinLength == 11:
        dayPart = CInt(Mid(cleanTIN, 1, 2))
        monthPart = CInt(Mid(cleanTIN, 3, 2))
        yearPart = CInt(Mid(cleanTIN, 5, 2))
        century = (monthPart - 1) // 20
        if century < 0 or century > 4:
            errorMsg.value = "Polonya TIN (Yapı 2) için ay değeri geçersiz."
            return _result
        if century == 0:
            fullYear = 1800 + yearPart
        elif century == 1:
            fullYear = 1900 + yearPart
        elif century == 2:
            fullYear = 2000 + yearPart
        elif century == 3:
            fullYear = 2100 + yearPart
        elif century == 4:
            fullYear = 2200 + yearPart
        birthdate = DateSerial(fullYear, (monthPart - 1) % 20 + 1, dayPart)
        if not IsDate(birthdate):
            errorMsg.value = "Polonya TIN'i (Yapı 2) için tarih bilgisi geçersizdir."
            return _result
        weights = Array(7, 3, 1, 9, 7, 3, 1, 7, 3)
        weightedSum = 0
        for i in vba_range(1, 9, 1):
            digit = CInt(Mid(cleanTIN, i, 1))
            weightedSum = weightedSum + digit * weights(i - 1)
        calculatedCheckDigit = weightedSum % 10
        calculatedCheckDigit = (10 - calculatedCheckDigit) % 10
        checkDigit = CInt(Right(cleanTIN, 1))
        if calculatedCheckDigit != checkDigit:
            errorMsg.value = "Geçersiz Polonya TIN (Yapı 2). Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
            return _result
        _result = True
    return _result

def ValidatePortugalTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    for i in vba_range(1, Len(tin), 1):
        if not IsNumeric(Mid(tin, i, 1)):
            errorMsg.value = "Portekiz TIN'i sadece rakamlardan oluşmalıdır. Harf veya özel karakterler içeremez."
            return _result
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) != 9:
        errorMsg.value = "Portekiz TIN'i 9 haneli olmalıdır."
        return _result
    weights = Array(9, 8, 7, 6, 5, 4, 3, 2)
    weightedSum = 0
    for i in vba_range(1, 8, 1):
        digit = CInt(Mid(cleanTIN, i, 1))
        weightedSum = weightedSum + digit * weights(i - 1)
    calculatedCheckDigit = 11 - (weightedSum % 11)
    if calculatedCheckDigit > 9:
        calculatedCheckDigit = 0
    checkDigit = CInt(Right(cleanTIN, 1))
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Portekiz TIN'i. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateQatarTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if Len(cleanTIN) == 10:
        if Left(cleanTIN, 1) == "5" and IsAllNumeric(cleanTIN):
            _result = True
        else:
            errorMsg.value = "GTA formatlı Katar TIN'i 5 ile başlamalı ve ardından 9 rakam gelmelidir."
    elif Len(cleanTIN) == 7:
        if Left(cleanTIN, 1) == "T":
            numericPart = Right(cleanTIN, 6)
            if IsAllNumeric(numericPart):
                _result = True
            else:
                errorMsg.value = "QFCA formatlı Katar TIN'i T ile başlamalı ve ardından 6 rakam gelmelidir."
        else:
            errorMsg.value = "QFCA formatlı Katar TIN'i T ile başlamalıdır."
    else:
        errorMsg.value = "Katar TIN'i GTA için 10 haneli (5 ile başlayan) veya QFCA için 7 haneli (T ile başlayan) olmalıdır."
    return _result

def ValidateRomaniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) == 13 and IsNumeric(tin):
        _result = True
    else:
        errorMsg.value = "Romanya TIN'i tam olarak 13 rakamdan oluşmalıdır. Başka karakter içermemelidir."
    return _result

def ValidateRussiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Rusya TIN'i sadece rakamlardan oluşmalıdır. Harf veya özel karakterler içeremez."
        return _result
    if Len(cleanTIN) == 10:
        _result = True
        return _result
    elif Len(cleanTIN) == 10 and Left(cleanTIN, 4) == "9909":
        _result = True
        return _result
    elif Len(cleanTIN) == 12:
        _result = True
        return _result
    else:
        errorMsg.value = "Rusya TIN'i geçersiz. 10 haneli (İşletmeler/Yabancı Kuruluşlar) veya 12 haneli (Bireyler) olmalıdır."
    return _result

def ValidateRwandaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Ruanda TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Ruanda TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateSaintLuciaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Trim(tin)
    if Len(cleanTIN) == 0:
        errorMsg.value = "Saint Lucia TIN'i boş olamaz."
        return _result
    if Len(cleanTIN) > 6:
        errorMsg.value = "Saint Lucia TIN'i en fazla 6 karakterden oluşmalıdır."
        return _result
    _result = True
    return _result

def ValidateSanMarinoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    prefix = ""
    numericPart = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) == 0:
        errorMsg.value = "San Marino TIN boş olamaz."
        return _result
    if IsAllNumeric(cleanTIN) and Len(cleanTIN) > 1:
        _result = True
        return _result
    if Len(cleanTIN) >= 2:
        prefix = Left(cleanTIN, 2)
        if prefix == "SM":
            numericPart = Right(cleanTIN, Len(cleanTIN) - 2)
            if Len(numericPart) == 5 and IsAllNumeric(numericPart):
                _result = True
                return _result
            else:
                errorMsg.value = "San Marino Tüzel Kişi TIN'i SM ile başlamalı ve ardından 5 rakam gelmelidir."
                return _result
    errorMsg.value = "Geçersiz San Marino TIN formatı. Bireyler için birden fazla rakam veya tüzel kişiler için SM + 5 rakam olmalı."
    return _result

def ValidateSaudiArabiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Suudi Arabistan TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10 and Len(cleanTIN) != 15:
        errorMsg.value = "Suudi Arabistan TIN'i 10 veya 15 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateSerbiaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Sırbistan TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Sırbistan TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateSeychellesTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Seyşeller TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Seyşeller TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateSingaporeTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Trim(tin)
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) == 0:
        errorMsg.value = "Singapur TIN boş olamaz."
        return _result
    if Left(tin, 1) == " ":
        errorMsg.value = "Singapur TIN'inde başında boşluk olamaz."
        return _result
    if (Len(cleanTIN) == 9 or Len(cleanTIN) == 10) and vba_like(cleanTIN, "*[A-Z0-9]*"):
        _result = True
    else:
        errorMsg.value = "Singapur TIN sadece 9 veya 10 alfasayısal karakter içermelidir."
        return _result
    return _result

def ValidateSintMaartenCRIB(crib, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanCRIB = ""
    cleanCRIB = Replace(Trim(crib), " ", "")
    if not IsAllNumeric(cleanCRIB):
        errorMsg.value = "Sint Maarten CRIB numarası sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanCRIB) != 11:
        errorMsg.value = "Sint Maarten CRIB numarası 11 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateSlovakiaTIN_Final(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    firstTwoDigits = 0
    monthPart = 0
    dayPart = 0
    isValid = False
    isValid = False
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "/", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Slovakya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) == 9 or Len(cleanTIN) == 10:
        firstTwoDigits = CInt(Left(cleanTIN, 2))
        if firstTwoDigits >= 0 and firstTwoDigits <= 99:
            if Len(cleanTIN) == 9:
                if firstTwoDigits >= 54:
                    errorMsg.value = "Slovakya TIN yapısı 1'de, ilk iki hane 54 veya daha büyükse, TIN 10 haneli olmalıdır."
                    return _result
            elif Len(cleanTIN) == 10:
                if firstTwoDigits < 54:
                    errorMsg.value = "Slovakya TIN yapısı 1'de, ilk iki hane 54'ten küçükse, TIN 9 haneli olmalıdır."
                    return _result
            monthPart = CInt(Mid(cleanTIN, 3, 2))
            if not ((monthPart >= 1 and monthPart <= 12) or (monthPart >= 51 and monthPart <= 62)):
                errorMsg.value = "Slovakya TIN'inde ay kısmı geçersiz (01-12 veya 51-62 olmalı)."
                return _result
            dayPart = CInt(Mid(cleanTIN, 5, 2))
            if dayPart < 1 or dayPart > 31:
                errorMsg.value = "Slovakya TIN'inde gün kısmı geçersiz (01-31 olmalı)."
                return _result
            isValid = True
            _result = True
        else:
            if Len(cleanTIN) == 10:
                isValid = True
                _result = True
        if not isValid:
            errorMsg.value = "Slovakya TIN'i geçerli bir yapıya sahip değil."
            return _result
    else:
        errorMsg.value = "Slovakya TIN'i 9 veya 10 haneli olmalıdır."
        return _result
    return _result

def ValidateSloveniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    digit = 0
    weightedSum = 0
    checkDigit = 0
    calculatedCheckDigit = 0
    weights = None
    digits = VbaArray(1, 8, 0)
    firstSevenDigits = 0
    remainder = 0
    if Len(tin) != 8:
        errorMsg.value = "Slovenya TIN'i tam olarak 8 rakamdan oluşmalıdır."
        return _result
    if not IsAllNumeric(tin):
        errorMsg.value = "Slovenya TIN'i sadece rakamlardan oluşmalıdır. Boşluk, harf veya özel karakter içeremez."
        return _result
    for i in vba_range(1, 8, 1):
        digits[i] = CInt(Mid(tin, i, 1))
    firstSevenDigits = CLng(Left(tin, 7))
    if firstSevenDigits < 1000000 or firstSevenDigits > 9999999:
        errorMsg.value = "Slovenya TIN'inin ilk yedi hanesi 1000000 ile 9999999 arasında olmalıdır."
        return _result
    weights = Array(0, 8, 7, 6, 5, 4, 3, 2)
    weightedSum = 0
    for i in vba_range(1, 7, 1):
        weightedSum = weightedSum + digits(i) * weights(i)
    remainder = weightedSum % 11
    if remainder == 0:
        calculatedCheckDigit = 0
    else:
        calculatedCheckDigit = 11 - remainder
        if calculatedCheckDigit == 10:
            calculatedCheckDigit = 0
    checkDigit = digits(8)
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz Slovenya TIN'i. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result

def ValidateSouthAfricaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    firstDigit = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Güney Afrika TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10:
        errorMsg.value = "Güney Afrika TIN'i 10 haneli olmalıdır."
        return _result
    firstDigit = Left(cleanTIN, 1)
    if firstDigit != "0" and firstDigit != "1" and firstDigit != "2" and firstDigit != "3" and firstDigit != "9":
        errorMsg.value = "Güney Afrika TIN'i 0, 1, 2, 3 veya 9 ile başlamalıdır."
        return _result
    _result = True
    return _result

def ValidateSouthKoreaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    yearOfBirth = 0
    sexDigit = ""
    cleanTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Güney Kore TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) == 13:
        yearOfBirth = CInt(Left(cleanTIN, 2))
        sexDigit = Mid(cleanTIN, 7, 1)
        if sexDigit == "1" or sexDigit == "2" or sexDigit == "5" or sexDigit == "6":
            if yearOfBirth >= 0 and yearOfBirth <= 21:
                yearOfBirth = 2000 + yearOfBirth
            else:
                yearOfBirth = 1900 + yearOfBirth
        elif sexDigit == "3" or sexDigit == "4" or sexDigit == "7" or sexDigit == "8":
            yearOfBirth = 2000 + yearOfBirth
        else:
            errorMsg.value = "Güney Kore TIN'inin 7. hanesi geçersiz cinsiyet kodudur."
            return _result
        _result = True
        return _result
    elif Len(cleanTIN) == 10:
        _result = True
        return _result
    elif Len(cleanTIN) == 13:
        _result = True
        return _result
    errorMsg.value = "Güney Kore TIN'i geçersiz. 13 haneli (gerçek kişiler) veya 10/13 haneli (tüzel kişiler) olmalıdır."
    return _result

def ValidateSpainVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    letters = None
    modResult = 0
    expectedLetter = ""
    actualLetter = ""
    firstChar = ""
    replacement = ""
    numericPart = ""
    numericPartNonNatural = ""
    letters = Array("", "T", "R", "W", "A", "G", "M", "Y", "F", "P", "D", "X", "B", "N", "J", "Z", "S", "Q", "V", "H", "L", "C", "K", "E")
    vkn = Trim(vkn)
    if Len(vkn) < 9:
        vkn = StringFunc(9 - Len(vkn), "0") + vkn
    if Len(vkn) != 9:
        errorMsg.value = "İspanya VKN'si 9 karakter olmalıdır."
        return _result
    firstChar = UCase(Left(vkn, 1))
    if vba_like(firstChar, "[A-Z]"):
        if firstChar == "X" or firstChar == "Y" or firstChar == "Z" or firstChar == "K" or firstChar == "L" or firstChar == "M":
            if firstChar == "X":
                replacement = "0"
            elif firstChar == "Y":
                replacement = "1"
            elif firstChar == "Z":
                replacement = "2"
            elif firstChar == "K" or firstChar == "L" or firstChar == "M":
                replacement = "0"
            numericPart = replacement + Mid(vkn, 2, 7)
            if not IsAllNumeric(numericPart):
                errorMsg.value = "İspanya VKN'si 2-8 karakterleri rakam olmalıdır."
                return _result
            modResult = CLng(numericPart) % 23
            modResult = modResult + 1
            expectedLetter = letters(modResult)
            actualLetter = UCase(Right(vkn, 1))
            if actualLetter == expectedLetter:
                _result = True
            else:
                errorMsg.value = "İspanya VKN'si geçersiz kontrol harfi. Beklenen: " + expectedLetter + ", Girilen: " + actualLetter + "."
        elif firstChar == "A" or firstChar == "B" or firstChar == "C" or firstChar == "D" or firstChar == "E" or firstChar == "F" or firstChar == "G" or firstChar == "H" or firstChar == "J" or firstChar == "P" or firstChar == "Q" or firstChar == "S" or firstChar == "U" or firstChar == "V" or firstChar == "N" or firstChar == "W":
            numericPartNonNatural = Mid(vkn, 2, 7)
            if not IsAllNumeric(numericPartNonNatural):
                errorMsg.value = "İspanya VKN'si 2-8 karakterleri rakam olmalıdır."
                return _result
            modResult = CLng(Mid(vkn, 2, 8)) % 23
            modResult = modResult + 1
            expectedLetter = letters(modResult)
            actualLetter = UCase(Right(vkn, 1))
            if actualLetter == expectedLetter:
                _result = True
            else:
                errorMsg.value = "İspanya VKN'si geçersiz kontrol harfi. Beklenen: " + expectedLetter + ", Girilen: " + actualLetter + "."
        else:
            errorMsg.value = "İspanya VKN'si geçerli bir başlangıç harfi ile başlamalıdır (A, B, C, D, E, F, G, H, J, P, Q, S, U, V, N, W, X, Y, Z, K, L, M)."
    else:
        if not IsAllNumeric(Left(vkn, 8)):
            errorMsg.value = "İspanya VKN'si ilk 8 karakter rakam olmalıdır."
            return _result
        modResult = CLng(Left(vkn, 8)) % 23
        modResult = modResult + 1
        expectedLetter = letters(modResult)
        actualLetter = UCase(Right(vkn, 1))
        if actualLetter == expectedLetter:
            _result = True
        else:
            errorMsg.value = "İspanya VKN'si geçersiz kontrol harfi. Beklenen: " + expectedLetter + ", Girilen: " + actualLetter + "."
    return _result

def ValidateSriLankaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Sri Lanka TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Sri Lanka TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateSwedenTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanedTIN = ""
    is12Digits = False
    isCoordinationNumber = False
    centuryPart = ""
    yearPart = ""
    monthPart = ""
    dayPart = ""
    checkDigit = 0
    calculatedCheckDigit = 0
    sum = 0
    i = 0
    digit = 0
    tempSum = 0
    currentYearVar = 0
    fullYear = 0
    cleanedTIN = Replace(Replace(Trim(tin), " ", ""), "-", "")
    if Len(cleanedTIN) == 10:
        is12Digits = False
    elif Len(cleanedTIN) == 12:
        is12Digits = True
    else:
        errorMsg.value = "İsveç TIN'i 10 veya 12 haneli sayı olmalıdır."
        return _result
    if not IsNumeric(cleanedTIN):
        errorMsg.value = "İsveç TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if is12Digits:
        centuryPart = Mid(cleanedTIN, 1, 2)
        yearPart = Mid(cleanedTIN, 3, 2)
        monthPart = Mid(cleanedTIN, 5, 2)
        dayPart = Mid(cleanedTIN, 7, 2)
        checkDigit = CInt(Right(cleanedTIN, 1))
    else:
        centuryPart = ""
        yearPart = Mid(cleanedTIN, 1, 2)
        monthPart = Mid(cleanedTIN, 3, 2)
        dayPart = Mid(cleanedTIN, 5, 2)
        checkDigit = CInt(Right(cleanedTIN, 1))
    if not IsNumeric(monthPart) or not IsNumeric(dayPart):
        errorMsg.value = "Ay ve gün değerleri sayısal olmalıdır."
        return _result
    if CInt(monthPart) < 1 or CInt(monthPart) > 12:
        errorMsg.value = "Ay değeri 01 ile 12 arasında olmalıdır."
        return _result
    if CInt(dayPart) >= 61 and CInt(dayPart) <= 91:
        isCoordinationNumber = True
        dayPart = Format(CInt(dayPart) - 60, "00")
    elif CInt(dayPart) >= 1 and CInt(dayPart) <= 31:
        isCoordinationNumber = False
    else:
        errorMsg.value = "Gün değeri 01 ile 31 veya 61 ile 91 arasında olmalıdır."
        return _result
    fullDate = None
    if is12Digits:
        fullYear = CInt(centuryPart + yearPart)
    else:
        currentYearVar = VBA.year(VBA.VBA.Date)
        currentCentury = 0
        currentCentury = (currentYearVar // 100) * 100
        fullYear = currentCentury + CInt(yearPart)
        if fullYear > currentYearVar:
            fullYear = fullYear - 100
    fullDate = DateSerial(fullYear, CInt(monthPart), CInt(dayPart))
    sum = 0
    digits = ""
    if is12Digits:
        digits = Mid(cleanedTIN, 3, 9)
    else:
        digits = Left(cleanedTIN, 9)
    for i in vba_range(1, Len(digits), 1):
        digit = CInt(Mid(digits, i, 1))
        if (i % 2) == 1:
            tempSum = digit * 2
            if tempSum > 9:
                tempSum = tempSum - 9
        else:
            tempSum = digit
        sum = sum + tempSum
    calculatedCheckDigit = (10 - (sum % 10)) % 10
    if calculatedCheckDigit != checkDigit:
        errorMsg.value = "Geçersiz İsveç TIN'i. Kontrol basamağı uyuşmuyor. Hesaplanan: " + CStr(calculatedCheckDigit) + ", Girilen: " + CStr(checkDigit)
        return _result
    _result = True
    return _result
    errorMsg.value = "Geçersiz tarih değeri."
    _result = False
    return _result

def ValidateSwitzerlandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if Left(cleanTIN, 3) == "756" or UCase(Left(cleanTIN, 3)) == "CHE":
        cleanTIN = Right(cleanTIN, Len(cleanTIN) - 3)
        if Len(cleanTIN) == 9 and IsAllNumeric(cleanTIN):
            _result = True
            return _result
        else:
            errorMsg.value = "İsviçre TIN'i '756' veya 'CHE' ön ekinden sonra 9 rakam içermelidir."
            return _result
    else:
        errorMsg.value = "İsviçre TIN'i '756' veya 'CHE' ön ekiyle başlamalıdır."
        return _result
    errorMsg.value = "İsviçre TIN'i geçersiz formatta."
    return _result

def ValidateSyriaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) > 16:
        errorMsg.value = "Suriye TIN'i en fazla 16 karakter olabilir."
        return _result
    _result = True
    return _result

def ValidateTaiwanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    firstTwoChars = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    cleanTIN = UCase(cleanTIN)
    if Len(cleanTIN) == 0:
        errorMsg.value = "Tayvan TIN boş olamaz."
        return _result
    if (Len(cleanTIN) == 10 and vba_like(Left(cleanTIN, 1), "[A-Z]") and vba_like(Mid(cleanTIN, 2, 9), "#########")):
        if IsNumeric(Mid(cleanTIN, 2, 9)):
            _result = True
            return _result
        else:
            errorMsg.value = "Tayvan Vatandaşlık Kartı Numarası formatı hatalı."
            return _result
    elif (Len(cleanTIN) == 10 and vba_like(Left(cleanTIN, 2), "[A-Z][A-Z]") and vba_like(Mid(cleanTIN, 3, 8), "########")):
        if IsNumeric(Mid(cleanTIN, 3, 8)):
            _result = True
            return _result
        else:
            errorMsg.value = "Tayvan Kimlik Numarası formatı hatalı."
            return _result
    elif (Len(cleanTIN) == 7 and Left(cleanTIN, 1) == "9" and vba_like(Mid(cleanTIN, 2, 6), "######")):
        if IsNumeric(Mid(cleanTIN, 2, 6)):
            _result = True
            return _result
        else:
            errorMsg.value = "Tayvan Vergi Numarası (Çin) formatı hatalı."
            return _result
    elif (Len(cleanTIN) == 10 and vba_like(Mid(cleanTIN, 1, 4), "####") and vba_like(Mid(cleanTIN, 5, 2), "##") and vba_like(Mid(cleanTIN, 7, 2), "[A-Z][A-Z]")):
        if IsNumeric(Left(cleanTIN, 4)) and IsNumeric(Mid(cleanTIN, 5, 2)) and IsAllLetters(Mid(cleanTIN, 7, 2)):
            _result = True
            return _result
        else:
            errorMsg.value = "Tayvan Vergi Numarası formatı hatalı."
            return _result
    elif (Len(cleanTIN) == 8 and IsAllNumeric(cleanTIN)):
        _result = True
        return _result
    else:
        errorMsg.value = "Geçersiz Tayvan TIN formatı."
        return _result
    return _result

def ValidateTajikistanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Tacikistan TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9:
        errorMsg.value = "Tacikistan TIN'i 9 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateTanzaniaTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Tanzanya TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10:
        errorMsg.value = "Tanzanya TIN'i 10 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateThailandTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Tayland TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 10 and Len(cleanTIN) != 13:
        errorMsg.value = "Tayland TIN'i 10 veya 13 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateTrinidadTobagoTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    if Len(tin) != 10:
        errorMsg.value = "Trinidad ve Tobago TIN'i tam olarak 10 rakamdan oluşmalıdır."
        return _result
    for i in vba_range(1, Len(tin), 1):
        if not vba_like(Mid(tin, i, 1), "#"):
            errorMsg.value = "Trinidad ve Tobago TIN'i sadece rakamlardan oluşmalıdır. Boşluk, harf veya özel karakter içeremez."
            return _result
    _result = True
    return _result

def ValidateTurkmenistanTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) > 16:
        errorMsg.value = "Türkmenistan TIN'i en fazla 16 karakter olabilir."
        return _result
    _result = True
    return _result

def ValidateUkraineTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    i = 0
    if Len(tin) == 8 or Len(tin) == 10:
        for i in vba_range(1, Len(tin), 1):
            if not vba_like(Mid(tin, i, 1), "#"):
                errorMsg.value = "Ukrayna TIN'i sadece rakamlardan oluşmalıdır. Boşluk, harf veya özel karakter içeremez."
                return _result
        _result = True
    else:
        errorMsg.value = "Ukrayna TIN'i 8 veya 10 haneli bir sayı olmalıdır."
    return _result

def ValidateUAEVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(vkn) == 15 and IsAllNumeric(vkn):
        _result = True
    else:
        errorMsg.value = "Birleşik Arap Emirlikleri VKN'si 15 karakter ve tamamen rakamlardan oluşmalıdır."
    return _result

def ValidateUKVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(vkn) == 10 and IsAllNumeric(vkn):
        _result = True
    elif Len(vkn) == 9:
        if vba_like(Mid(vkn, 1, 1), "[A-Za-z]") and vba_like(Mid(vkn, 2, 1), "[A-Za-z]") and IsAllNumeric(Mid(vkn, 3, 6)) and vba_like(Mid(vkn, 9, 1), "[A-Za-z]"):
            _result = True
        else:
            errorMsg.value = "Birleşik Krallık VKN'si alfanümerik ve 9 karakter uzunluğunda olmalıdır."
    else:
        errorMsg.value = "Birleşik Krallık VKN'si ya tamamen rakamlardan oluşmalı ya da belirtilen alfanümerik formata uymalıdır."
    return _result

def ValidateUSAVKN(vkn, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    regexSSN = None
    regexEIN = None
    regexPlain9Digits = None
    regexSSN = CreateObject("VBScript.RegExp")
    regexSSN.Pattern = "^\d{3}-\d{2}-\d{4}$"
    regexSSN.IgnoreCase = False
    regexSSN.Global = False
    regexEIN = CreateObject("VBScript.RegExp")
    regexEIN.Pattern = "^\d{2}-\d{7}$"
    regexEIN.IgnoreCase = False
    regexEIN.Global = False
    regexPlain9Digits = CreateObject("VBScript.RegExp")
    regexPlain9Digits.Pattern = "^\d{9}$"
    regexPlain9Digits.IgnoreCase = False
    regexPlain9Digits.Global = False
    if regexSSN.Test(vkn) or regexEIN.Test(vkn) or regexPlain9Digits.Test(vkn):
        _result = True
    else:
        errorMsg.value = "ABD VKN'si xxx-xx-xxxx, xx-xxxxxxx veya 9 haneli sayı formatında olmalıdır."
    return _result

def ValidateUruguayTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Uruguay TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 9 and Len(cleanTIN) != 12:
        errorMsg.value = "Uruguay TIN'i 9 veya 12 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateVanuatuTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    cleanTIN = Replace(Trim(tin), " ", "")
    if not IsAllNumeric(cleanTIN):
        errorMsg.value = "Vanuatu TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(cleanTIN) != 6:
        errorMsg.value = "Vanuatu TIN'i 6 haneli olmalıdır."
        return _result
    _result = True
    return _result

def ValidateVietnamTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    cleanTIN = ""
    hyphenCount = 0
    i = 0
    char = ""
    tempTIN = ""
    cleanTIN = Trim(tin)
    if Len(cleanTIN) == 0:
        errorMsg.value = "Vietnam TIN boş olamaz."
        return _result
    tempTIN = ""
    hyphenCount = 0
    for i in vba_range(1, Len(cleanTIN), 1):
        char = Mid(cleanTIN, i, 1)
        if char == "-":
            hyphenCount = hyphenCount + 1
        else:
            tempTIN = tempTIN + char
    if hyphenCount > 1:
        errorMsg.value = "Vietnam TIN'i en fazla 1 tire (-) içerebilir."
        return _result
    if not IsAllNumeric(tempTIN):
        errorMsg.value = "Vietnam TIN'i sadece rakamlardan oluşmalıdır."
        return _result
    if Len(tempTIN) == 10:
        _result = True
        return _result
    elif Len(tempTIN) == 13 and hyphenCount <= 1:
        _result = True
        return _result
    else:
        errorMsg.value = "Vietnam TIN'i 10 haneli (tire olmadan) veya 13 haneli (tek tire ile) olmalıdır."
        return _result
    return _result

def ValidateYemenTIN(tin, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    if Len(tin) > 16:
        errorMsg.value = "Yemen TIN'i en fazla 16 karakter olabilir."
        return _result
    _result = True
    return _result

def ValidateCountryDispatch(vkn, countryCode, formatInfo, errorMsg):
    _result = False
    errorMsg = ensure_ref(errorMsg)
    formatInfo = ensure_ref(formatInfo)
    isValid = False
    isValid = False
    formatInfo.value = ""
    errorMsg.value = ""
    if countryCode == "AD":
        formatInfo.value = "8 karakter, belirli harf ve sayısal aralıklarla"
        if ValidateAndorraVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AE":
        formatInfo.value = "15 rakam, tamamen rakam"
        if ValidateUAEVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AF":
        formatInfo.value = "10 haneli sayı. Sadece rakamlardan oluşur."
        if ValidateAfghanistanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AI":
        formatInfo.value = "10 rakam, bireyler için 1 ile, işletmeler için 2 ile başlar."
        if ValidateAnguillaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AL":
        formatInfo.value = "10 karakter, alfanümerik"
        if ValidateAlbaniaVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AM":
        formatInfo.value = "8 haneli sayı"
        if ValidateArmeniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AR":
        formatInfo.value = "11 haneli sayı, belirli ön ekler ile (20, 23, 24, 27, 30, 33)"
        if ValidateArgentinaCUIT(vkn, errorMsg):
            isValid = True
    elif countryCode == "AT":
        formatInfo.value = "9 haneli sayı, kontrol basamağı ile"
        if ValidateAustriaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AU":
        formatInfo.value = "8, 9 veya 11 haneli sayı (boşluklar hariç)"
        if ValidateAustraliaVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AW":
        formatInfo.value = "8 haneli sayı"
        if ValidateArubaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "AZ":
        formatInfo.value = "10 haneli (tamamen rakam) veya 7 haneli (harf ve rakam)"
        if ValidateAzerbaijanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BA":
        formatInfo.value = "12 haneli sayı, sadece rakamlardan oluşur"
        if ValidateBosniaHerzegovinaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BB":
        formatInfo.value = "13-digit number starting with '1', no letters or symbols."
        if ValidateBarbadosTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BE":
        formatInfo.value = "11 haneli sayı, kontrol basamaklarıyla (son iki hane)"
        if ValidateBelgiumTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BF":
        formatInfo.value = "8 rakam ve ardından bir harf, tamamen rakamlardan oluşur."
        if ValidateBurkinaFasoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BG":
        formatInfo.value = "10 haneli sayı, kontrol basamağı ile"
        if ValidateBulgariaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BH":
        formatInfo.value = "15 rakam, tamamen rakam"
        if ValidateBahrainTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BI":
        formatInfo.value = "10 haneli sayı, sadece rakamlardan oluşur."
        if ValidateBurundiTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BN":
        formatInfo.value = "9, 10 veya 11 karakter. Bir tane tire (-) içerebilir."
        if ValidateBruneiTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BO":
        formatInfo.value = "7 veya 10 haneli sayı, sadece rakamlardan oluşur."
        if ValidateBoliviaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BR":
        formatInfo.value = "11 haneli (CPF - gerçek kişi) veya 14 haneli (CNPJ - tüzel kişi)"
        if ValidateBrazilCPF_CNPJ(vkn, errorMsg):
            isValid = True
    elif countryCode == "BT":
        formatInfo.value = "AAA##### formatında (harf + harf + harf + rakam + rakam + rakam + rakam + rakam)"
        if ValidateBhutanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BW":
        formatInfo.value = "Bir harf ile başlayan ve ardından 9 veya 10 rakam gelen TIN."
        if ValidateBotswanaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BY":
        formatInfo.value = "9 karakter (alfanümerik)"
        if ValidateBelarusTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "BZ":
        formatInfo.value = "6 rakam, tamamen rakam"
        if ValidateBelizeTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CA":
        formatInfo.value = "9 haneli SIN/BN veya 'T' + 8 haneli Trust Hesap Numarası"
        if ValidateCanadaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CD":
        formatInfo.value = "7, 8 veya 9 karakter uzunluğunda, alfanümerik."
        if ValidateDemocraticRepublicOfCongoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CF":
        formatInfo.value = "7 rakam ve ardından bir harf."
        if ValidateCentralAfricanRepublicTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CH":
        formatInfo.value = "'756' veya 'CHE' ön eki ve ardından 9 rakam"
        if ValidateSwitzerlandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CK":
        formatInfo.value = "5 rakam, tamamen rakam"
        if ValidateCookIslandsTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CL":
        formatInfo.value = "xx.xxx.xxx-x formatında, Modulo 11 kontrolü ile"
        if ValidateChileTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CM":
        formatInfo.value = "Bir harf, 12 rakam ve bir harf kombinasyonu (A#########A) formatında 14 karakter."
        if ValidateCameroonNIU(vkn, errorMsg):
            isValid = True
    elif countryCode == "CN":
        formatInfo.value = "15 veya 18 karakter uzunluğunda olabilir. Rakamlar ve harflerden oluşur. Boşluk veya özel karakter içeremez."
        if ValidateChinaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CO":
        formatInfo.value = "Ana numara 1-13 basamak + kontrol rakamı. Farklı türler için belirli aralıklar."
        if ValidateColombiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CR":
        formatInfo.value = "Bireysel: 9 rakam, Kurumsal/NITE: 10 rakam, DIMEX: 11-12 rakam"
        if ValidateCostaRicaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CW":
        formatInfo.value = "9 rakam, sadece rakamlardan oluşur"
        if ValidateCuracaoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CY":
        formatInfo.value = "8 rakam + 1 büyük harf (toplam 9 karakter)"
        if ValidateCyprusTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "CZ":
        formatInfo.value = "9 veya 10 haneli sayı"
        if ValidateCzechiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "DE":
        formatInfo.value = "11 karakter, tamamen rakam, 0 ile başlamaz. TCKN numarası ile aynı olmaması gerekir."
        if ValidateGermanyVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "DK":
        formatInfo.value = "CPR (10 haneli) veya CVR/SE (8 haneli)"
        if ValidateDenmarkTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "DM":
        formatInfo.value = "6 rakam (vergi mükellef numarası) veya 7 rakam (TIN), tamamen rakam"
        if ValidateDominicaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "DZ":
        formatInfo.value = "15 veya 20 haneli sayı, sadece rakamlardan oluşur."
        if ValidateAlgeriaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "EC":
        formatInfo.value = "13 rakam, il kodu, kişi türü ve kontrol rakamı doğrulaması ile"
        if ValidateEcuadorRUC(vkn, errorMsg):
            isValid = True
    elif countryCode == "EE":
        formatInfo.value = "11 haneli sayı, kontrol basamağı ile"
        if ValidateEstoniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "EG":
        formatInfo.value = "9 haneli sayı, isteğe bağlı tirelerle"
        if ValidateEgyptTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "ES":
        formatInfo.value = "9 karakter, belirli formatlar ve check digit doğrulaması"
        if ValidateSpainVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "FI":
        formatInfo.value = "11 karakter, GGAAyy-KKKT formatında (K kontrol basamağı)"
        if ValidateFinlandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "FO":
        formatInfo.value = "9 rakam, ddmmyyxxx formatında, ilk 6 rakam geçerli bir tarih olmalı"
        if ValidateFaroeIslandsTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "FR":
        formatInfo.value = "13 rakam (birey) veya 9 rakam (kuruluş)"
        if ValidateFranceVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GB":
        formatInfo.value = "10 karakter, tamamen rakam veya 9 karakter, alfanümerik"
        if ValidateUKVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GD":
        formatInfo.value = "Tam olarak 6 rakamdan oluşmalıdır."
        if ValidateGrenadaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GE":
        formatInfo.value = "7 haneli sayı"
        if ValidateGeorgiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GH":
        formatInfo.value = "GRA TIN (11 karakter) veya NIA Ghanacard PIN (15 karakter)"
        if ValidateGhanaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GI":
        formatInfo.value = "Sadece rakamlardan oluşmalıdır."
        if ValidateGibraltarTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GL":
        formatInfo.value = "10 haneli (gerçek kişiler) veya 8 haneli (tüzel kişiler)"
        if ValidateGreenlandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GR":
        formatInfo.value = "9 haneli sayı"
        if ValidateGreeceTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "GT":
        formatInfo.value = "7 veya 8 haneli sayı, sadece rakamlardan oluşur."
        if ValidateGuatemalaNIT(vkn, errorMsg):
            isValid = True
    elif countryCode == "HK":
        formatInfo.value = "Bireyler için HKID numarası (1-2 harf + 6 rakam + 1 kontrol karakteri), kurumlar için BR numarası (8 rakam)"
        if ValidateHongKongTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "HR":
        formatInfo.value = "11 haneli sayı. İlk 10 hane rastgele sayı, son hane kontrol basamağıdır."
        if ValidateCroatiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "HT":
        formatInfo.value = "10 haneli sayı, sadece rakamlardan oluşur."
        if ValidateHaitiTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "HU":
        formatInfo.value = "10 haneli sayı. İlk hane 8 olmalı. Son hane kontrol basamağıdır."
        if ValidateHungaryTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "ID":
        formatInfo.value = "15 veya 16 rakam, sadece rakamlardan oluşmalıdır."
        if ValidateIndonesiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IE":
        formatInfo.value = "7 rakam + 1 veya 2 harf (toplam 8 veya 9 karakter)"
        if ValidateIrelandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IL":
        formatInfo.value = "Tam olarak 9 rakam, sadece rakamlardan oluşmalıdır."
        if ValidateIsraelTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IM":
        formatInfo.value = "Vergi Referans Numarası: [Ön Ek][6 haneli sayı][İsteğe bağlı - 2 haneli ek]; Örnekler: H123456, C654321-12" + vbCrLf + "Ulusal Sigorta Numarası: [2 harf][6 haneli sayı][A, B, C veya D]; Örnek: MA123456A"
        if ValidateIsleOfManTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IN":
        formatInfo.value = "10 karakter, ilk üç ve 5, 10. karakterler harf, 4. karakter özel durum kodu, 6-9 rakam."
        if ValidateIndiaPAN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IQ":
        formatInfo.value = "9 haneli (tüzel kişiler) veya 10 haneli (gerçek kişiler), sadece rakamlardan oluşur."
        if ValidateIraqTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IS":
        formatInfo.value = "10 rakam, ilk 6 rakam GGAAYY formatında geçerli bir tarih olmalı, 10. rakam yüzyılı belirtir (9, 0 veya 1)"
        if ValidateIcelandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "IT":
        formatInfo.value = "Codice Fiscale: 16 karakter, alfanümerik; Partita IVA: 11 karakter, tamamen rakam"
        if ValidateItalyVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "JE":
        formatInfo.value = "10 haneli sayı, opsiyonel olarak XXX-XXX-XXXX formatında tireler içerebilir."
        if ValidateJerseyTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "JM":
        formatInfo.value = "9 rakam, ilk rakam 0 veya 1 olmalıdır."
        if ValidateJamaicaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "JP":
        formatInfo.value = "12 haneli (Bireysel) veya 13 haneli (Kurumsal), sadece rakamlar"
        if ValidateJapanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "KE":
        formatInfo.value = "11 karakter (ilk ve son harf, diğerleri rakam), 'P' (tüzel) veya 'A' (gerçek) ile başlar"
        if ValidateKenyaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "KH":
        formatInfo.value = "A##-######### formatında (1 harf, 2 rakam, tire, 9 rakam)"
        if ValidateCambodiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "KR":
        formatInfo.value = "Bireyler için 13 haneli, Tüzel kişiler için 10 veya 13 haneli"
        if ValidateSouthKoreaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "KW":
        formatInfo.value = "Bireyler için: 12 haneli Civil ID numarası (C YYMMDD P SSSS formatında)." + vbCrLf + "Tüzel kişiler için: 6 haneli sayı." + vbCrLf + "TIN sadece rakamlardan oluşmalıdır. Harf, boşluk veya özel karakter içeremez."
        if ValidateKuwaitTIN(vkn, errorMsg):
            isValid = True
        else:
            isValid = False
    elif countryCode == "KZ":
        formatInfo.value = "IIN (12 haneli) veya BIN (12 haneli)"
        if Len(vkn) == 12 and IsAllNumeric(vkn):
            if ValidateKazakhstanIIN(vkn, errorMsg):
                isValid = True
            elif ValidateKazakhstanBIN(vkn, errorMsg):
                isValid = True
        else:
            errorMsg.value = "Kazakistan TIN'i 12 haneli ve sayısal olmalıdır."
    elif countryCode == "LB":
        formatInfo.value = "En fazla 16 karakter"
        if ValidateLebanonTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LC":
        formatInfo.value = "En fazla 6 karakterden oluşmalıdır."
        if ValidateSaintLuciaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LI":
        formatInfo.value = "1 ile 12 arasında rakam"
        if ValidateLiechtensteinTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LK":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateSriLankaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LS":
        formatInfo.value = "9 veya 10 karakter uzunluğunda olmalıdır."
        if ValidateLesothoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LT":
        formatInfo.value = "11 haneli sayı, kontrol basamağı ile"
        if ValidateLithuaniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LV":
        formatInfo.value = "11 haneli sayı (yapı 1 veya 2)"
        if ValidateLatviaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "LY":
        formatInfo.value = "En fazla 16 karakter (serbest metin)."
        if ValidateLibyaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MA":
        formatInfo.value = "15 haneli sayı, sadece rakamlardan oluşur."
        if ValidateMoroccoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MC":
        formatInfo.value = "14 haneli sayı, sadece tüzel kişiler için. Bireysel vergi numarası bulunmamaktadır."
        if ValidateMonacoSIRET(vkn, errorMsg):
            isValid = True
    elif countryCode == "MD":
        formatInfo.value = "7, 8 veya 13 haneli sayı, sadece rakamlardan oluşur."
        if ValidateMoldovaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "ME":
        formatInfo.value = "8 veya 13 haneli sayı, sadece rakamlardan oluşur."
        if ValidateMontenegroTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MG":
        formatInfo.value = "En fazla 16 karakter (serbest metin)."
        if ValidateMadagascarTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MH":
        formatInfo.value = "EIN: #####-04 (5 rakam, tire, 04) veya Employee No: 04-###### (04, tire, 6 rakam)."
        if ValidateMarshallIslandsTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MK":
        formatInfo.value = "13 haneli sayı veya 'MK' ile başlayıp ardından 13 haneli sayı olmalı, sadece rakamlardan oluşur."
        if ValidateNorthMacedoniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MN":
        formatInfo.value = "7 haneli sayı, sadece rakamlardan oluşur."
        if ValidateMongoliaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MO":
        formatInfo.value = "8 haneli sayı, sadece rakamlardan oluşur."
        if ValidateMacaoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MT":
        formatInfo.value = "Malta 1: 8 karakter (C1-C7 rakam, C8 harf); Malta 2: 9 rakam, ilk iki hanesi {11,22,33,44,55,66,77,88}"
        if ValidateMaltaTIN_Final(vkn, errorMsg):
            isValid = True
    elif countryCode == "MU":
        formatInfo.value = "Individuals: 8 rakam, ilk rakam {1,5,7,8}. Entities: 8 rakam, ilk rakam {2,3}."
        if ValidateMauritiusTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MV":
        formatInfo.value = "7 haneli rakam (BP) + 2-5 harf (Revenue Code) + 3 rakam (Sıra No). Örn: 1000001GST501"
        if ValidateMaldivesTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MX":
        formatInfo.value = "Bireyler için 13 karakter (4 harf + 6 rakam + 3 harf/rakam), Tüzel kişiler için 12 karakter (3 harf + 6 rakam + 3 harf/rakam)"
        if ValidateMexicoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MY":
        formatInfo.value = "Bireysel: IG ile başlayıp ardından rakamlar. Non-Individual: Kod+Rakamlar+'0'. Yoksa NRIC (12 hane rakam)."
        if ValidateMalaysiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "MZ":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateMozambiqueTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NA":
        formatInfo.value = "8 haneli sayı, sadece rakamlardan oluşur."
        if ValidateNamibiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NG":
        formatInfo.value = "FIRS TIN (8 hane + '-0001') veya JTB TIN (10 hane)"
        if ValidateNigeriaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NL":
        formatInfo.value = "9 haneli sayı, kontrol basamağı ile."
        if ValidateNetherlandsTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NO":
        formatInfo.value = "Ulusal Kimlik Numarası (11 hane) veya Organizasyon Numarası (9 hane) veya KDV Kayıtlı Tüzel Kişi (9 hane + 'MVA')"
        if ValidateNorwayTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NP":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateNepalTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NR":
        formatInfo.value = "9 haneli benzersiz sayı"
        if ValidateNauruTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "NZ":
        formatInfo.value = "8 veya 9 haneli sayı, kontrol basamağı ile (GG-GGG-GGG veya GGG-GGG-GGG)"
        if ValidateNewZealandIRD(vkn, errorMsg):
            isValid = True
    elif countryCode == "OM":
        formatInfo.value = "En fazla 7 haneli sayı"
        if ValidateOmanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "PE":
        formatInfo.value = "11 haneli sayı. Boşluk, harf veya özel karakter içeremez."
        if ValidatePeruTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "PH":
        formatInfo.value = "9-12 haneli, sadece rakamlardan oluşur, bireyler 1-9 ile, kuruluşlar 0 ile başlamalıdır."
        if ValidatePhilippinesTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "PK":
        formatInfo.value = "7 veya 13 haneli sayı. Boşluk, harf veya özel karakter içeremez."
        if ValidatePakistanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "PL":
        formatInfo.value = "10 haneli (Yapı 1) veya 11 haneli (Yapı 2)"
        if ValidatePolandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "PT":
        formatInfo.value = "9 haneli sayı, kontrol basamağı ile"
        if ValidatePortugalTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "QA":
        formatInfo.value = "GTA için 10 haneli (5 ile başlayan) veya QFCA için 7 haneli ('T' + 6 hane)"
        if ValidateQatarTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "RO":
        formatInfo.value = "13 haneli sayı"
        if ValidateRomaniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "RS":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateSerbiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "RU":
        formatInfo.value = "10 haneli (İşletmeler/Yabancı Kuruluşlar) veya 12 haneli (Bireyler)"
        if ValidateRussiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "RW":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateRwandaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SA":
        formatInfo.value = "10 veya 15 haneli sayı. Sadece rakamlardan oluşur"
        if ValidateSaudiArabiaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SC":
        formatInfo.value = "9 haneli sayı, kontrol basamağı ile."
        if ValidateSeychellesTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SE":
        formatInfo.value = "İsveç TIN'i 11 veya 13 karakterden oluşur (hyphen dahil). Hyphen zorunludur ve doğru pozisyonda yer alır." + vbCrLf + "Sweden 1 ve 2: YYMMDD-XXXX formatında (hyphen 7. pozisyonda)." + vbCrLf + "Sweden 3 ve 4: CCYYMMDD-XXXX formatında (hyphen 9. pozisyonda)." + vbCrLf + "Sweden 2 ve 4'te, gün değeri 60 eklenmiş olarak verilir (61-91)."
        if ValidateSwedenTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SG":
        formatInfo.value = "Değişken uzunluk ve formatlarda, özel karakterler içerir."
        if ValidateSingaporeTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SI":
        formatInfo.value = "8 haneli sayı"
        if ValidateSloveniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SK":
        formatInfo.value = "Slovakya TIN'i 9 veya 10 haneli olmalıdır (yapıya göre)"
        if ValidateSlovakiaTIN_Final(vkn, errorMsg):
            isValid = True
    elif countryCode == "SM":
        formatInfo.value = "Bireyler için birden fazla rakam, tüzel kişiler için 'SM' + 5 rakam"
        if ValidateSanMarinoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "SX":
        formatInfo.value = "11 haneli sayı"
        if ValidateSintMaartenCRIB(vkn, errorMsg):
            isValid = True
    elif countryCode == "SY":
        formatInfo.value = "En fazla 16 karakter (serbest metin)."
        if ValidateSyriaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "TH":
        formatInfo.value = "13 haneli sayı, sadece rakamlardan oluşur."
        if ValidateThailandTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "TJ":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateTajikistanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "TM":
        formatInfo.value = "En fazla 16 karakter (serbest metin)"
        if ValidateTurkmenistanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "TT":
        formatInfo.value = "10 haneli sayı. Sadece rakamlardan oluşur. Boşluk, harf veya özel karakter içeremez."
        if ValidateTrinidadTobagoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "TW":
        formatInfo.value = "Çeşitli formatlarda (10 haneli Ulusal Kimlik Kartı, 10 haneli Kimlik Numarası, 7 veya 10 haneli Vergi Kodu, 8 haneli BAN) alfanümerik veya numerik."
        if ValidateTaiwanTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "TZ":
        formatInfo.value = "10 haneli sayı, sadece rakamlardan oluşur."
        if ValidateTanzaniaTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "UA":
        formatInfo.value = "8 veya 10 haneli sayı. Sadece rakamlardan oluşur. Boşluk, harf veya özel karakter içeremez."
        if ValidateUkraineTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "US":
        formatInfo.value = "xxx-xx-xxxx, xx-xxxxxxx veya 9 haneli sayı formatında, tamamen rakam"
        if ValidateUSAVKN(vkn, errorMsg):
            isValid = True
    elif countryCode == "UY":
        formatInfo.value = "9 veya 12 haneli sayı, sadece rakamlardan oluşur."
        if ValidateUruguayTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "VN":
        formatInfo.value = "10 haneli sayı veya 13 haneli (1 adet tire içerebilir), sadece rakamlardan oluşur."
        if ValidateVietnamTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "VU":
        formatInfo.value = "6 haneli sayı. Sadece rakamlardan oluşur."
        if ValidateVanuatuTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "XK":
        formatInfo.value = "9 haneli sayı, sadece rakamlardan oluşur."
        if ValidateKosovoTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "YE":
        formatInfo.value = "En fazla 16 karakter (serbest metin)"
        if ValidateYemenTIN(vkn, errorMsg):
            isValid = True
    elif countryCode == "ZA":
        formatInfo.value = "10 haneli sayı. 0, 1, 2, 3 veya 9 ile başlar."
        if ValidateSouthAfricaTIN(vkn, errorMsg):
            isValid = True
    else:
        isValid = False
        formatInfo.value = "Desteklenmeyen ülke kodu: " + countryCode
        errorMsg.value = "Ülke kodu tanınmıyor veya desteklenmiyor."
    _result = isValid
    return _result
