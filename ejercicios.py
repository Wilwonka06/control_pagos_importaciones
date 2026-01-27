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

#serie fibonacci
def fibonacci(num):
    num+=1
    if num == 0:
        return 0
    elif num == 1:
        return 1
    else:
        return fibonacci(num - 1) + fibonacci(num - 2)
print(fibonacci(n))

