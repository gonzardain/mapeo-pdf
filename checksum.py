# -*- coding: utf-8 -*-
_patterns = {
    'A':    (10),      'B': (11),
    'C':    (12),      'D': (13),
    'E':    (14),      'F': (15),
    'G':    (16),      'H': (17),
    'I':    (18),      'J': (19),
    'K':    (20),      'L': (21),
    'M':    (22),      'N': (23),
    'O':    (24),      'P': (25),
    'Q':    (26),      'R': (27),
    'S':    (28),      'T': (29),
    'U':    (30),      'V': (31),
    'W':    (32),      'X': (33),
    'Y':    (34),      'Z': (35),
    }


def add_check_digit(code39):
    """
    Devuelve un dígito verificador a partir de una serie de digitos compuestos
    por 21 carácteres de la siguiente forma:

    Tipo de Código 3 de 9 Full ASCII (USS 3 de 9)
    Clave de Cliente: 3 dígitos (CYF)
    Clave del producto: 2 Dígitos (29)
    Numero de identificación única: 16 Dígitos (Linea de captura de predial
    pos [] del array)

    La presente función retorna un valor str(check_digit) que corresponde a
    Digito verificador 1 Dígito

    Total de digitos 22

    Nota importante debido a que la cadena a codificar contiene letras el
    presente codigo ha sido manipulado para que solo funcione con las letras
    correspondientes a CYF, se deja comentado el desarrollo a implementar para
    que funcione con las letras del alfabeto americano.
    """

    code39 = str(code39)
    if len(code39) != 18:
        raise Exception("Invalid length")

    odd_sum = 0
    even_sum = 0
    for i, char in enumerate(code39):

        j = i+1
        if j % 2 == 0:
            even_sum += int(char)

        else:
            odd_sum += int(char)

    """
    Para el correcto funcionamiento regresar a
    total_sum = (odd_sum * 3) + even_sum
    y codificar el array _patterns

    """
    total_sum = (even_sum * 3) + odd_sum +54
    print total_sum
    mod = total_sum % 10
    check_digit = 10 - mod
    if check_digit == 10:
        check_digit = 0
    #print code39 + str(check_digit)


    return (str(check_digit))


#if __name__ == "__main__":
    #code39 = "Code goes here"
    #add_check_digit(code39)
