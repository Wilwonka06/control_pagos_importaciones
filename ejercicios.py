#1 Fizz Buzz

num = 0

""" for num in range(101):
    if num % 3 == 0:
        print("Fizz")
    elif num % 5 == 0:
        print("Buzz")
    elif num % 3 == 0 and num % 5 == 0:
        print("FizzBuzz")
    else:
        print(num) """

Palabra1 = "Hola Mundo"
Palabra2 = "Mundo Hola"

""" def es_anagrama(palabra1, palabra2):
    if sorted(palabra1.upper()) == sorted(palabra2.upper()):
        return True
    else:
        return False

print(es_anagrama(Palabra1, Palabra2)) """


""" #successionfibonacci

def fibonacci():
    for i in range(50):
        print(i)
        if i == 0:
            i+1
        elif i == 1:
            i+1
        else:
            print(i-1 + i-2)

fibonacci() """


for i in range(101):
    if i % 2 != 0:
        print(i)
