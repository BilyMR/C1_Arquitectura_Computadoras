# Simulador de CPU de 8 Bits (Arquitectura Von Neumann) en Excel / VBA

**Universidad Católica Boliviana "San Pablo"**

**Unidad Académica Regional Santa Cruz**

**Departamento de Ingenierías y Ciencias Exactas**

**Carrera:** Ingeniería de Software

**Materia:** Arquitectura de Computadoras (SIS-131)

**Semestre:** 1/2026

**Evaluación:** Parcial 1 — Proyecto Práctico y Defensa Oral

---

## 1. Descripción General del Proyecto

Este proyecto consiste en la implementación de un **simulador interactivo y visual a nivel de transferencia de registros (RTL)** de una CPU de 8 bits basada en la **Arquitectura Von Neumann**. El sistema está completamente desarrollado en **Microsoft Excel** utilizando **VBA (Visual Basic for Applications)** como motor lógico y de animación gráfica.

El simulador modela con rigor técnico el ciclo completo de instrucción (**Fetch - Decode - Execute - Store**), la gestión de memoria RAM de 256 bytes mapeada en formato hexadecimal, la interacción entre la Unidad de Control y la ALU, y la animación en tiempo real del transporte de datos a través de los buses del sistema.

---

## 2. Especificaciones Técnicas del Sistema

### 2.1. Arquitectura de Memoria RAM (256 Bytes: `00h` - `FFh`)

La memoria principal está organizada en una matriz de 16 x 16 celdas de 8 bits, segmentada de la siguiente manera:

- **Segmento de Código (`00h` a `0Fh` - 16 Bytes):** Espacio reservado para los opcodes y operandos del programa en ensamblador.
- **Segmento de Datos (`10h` a `EFh` - 224 Bytes):** Espacio para lectura y escritura de variables y resultados del procesamiento.
- **Segmento de Pila / Stack (`F0h` a `FFh` - 16 Bytes):** Estructura LIFO utilizada para almacenamiento temporal, con el puntero de pila (**SP**) inicializado por defecto en `FFh`.

### 2.2. Registros Internos del Procesador

- **PC (Program Counter):** Registro de 8 bits que almacena la dirección de memoria de la siguiente instrucción a ejecutar.
- **IR (Instruction Register):** Guarda el Opcode de la instrucción leída durante la fase Fetch.
- **MAR (Memory Address Register):** Mantiene la dirección de RAM sobre la cual se ejecutará un ciclo de lectura o escritura.
- **MDR / MBR (Memory Data Register):** Almacena el dato de 8 bits que viaja desde o hacia la memoria RAM.
- **AX (Acumulador):** Registro multipropósito de 8 bits para resultados aritméticos y lógicos.
- **BX (Registro Base):** Registro multipropósito de 8 bits, utilizado frecuentemente como contador en bucles.
- **SP (Stack Pointer):** Puntero de 8 bits que apunta al tope de la pila (`FFh`).

### 2.3. Banderas de Estado (Flags Register)

- **ZF (Zero Flag):** Se establece en `1` si el resultado de la última operación de la ALU es `00h`; de lo contrario es `0`.
- **CF (Carry Flag):** Se establece en `1` si ocurre un desbordamiento sin signo (> 255).
- **SF (Sign Flag):** Se establece en `1` si el resultado tiene el bit más significativo en `1` (valor negativo en complemento a dos).

---

## 3. Conjunto de Instrucciones (ISA)

El procesador implementa una arquitectura CISC simplificada con codificación de longitud variable (1 o 2 bytes):

| Opcode (Hex) | Mnemónico / Formato | Tamaño | Descripción Operacional |
| :---: | :--- | :---: | :--- |
| **`10h`** | `MOV AX, imm` | 2 Bytes | Carga el valor inmediato (`imm`) en el registro `AX`. |
| **`11h`** | `MOV BX, imm` | 2 Bytes | Carga el valor inmediato (`imm`) en el registro `BX`. |
| **`13h`** | `STORE [dir], AX` | 2 Bytes | Almacena el contenido del registro `AX` en la dirección de RAM `[dir]`. |
| **`14h`** | `STORE [dir], BX` | 2 Bytes | Almacena el contenido del registro `BX` en la dirección de RAM `[dir]`. |
| **`20h`** | `ADD AX, imm` | 2 Bytes | Suma el valor inmediato (`imm`) al registro `AX` (`AX = AX + imm`). |
| **`21h`** | `SUB AX, imm` | 2 Bytes | Resta el valor inmediato (`imm`) al registro `AX` (`AX = AX - imm`). |
| **`23h`** | `DEC BX` | 1 Byte | Decrementa en `1` el registro `BX` (`BX = BX - 1`) y actualiza la bandera `ZF`. |
| **`25h`** | `ADD AX, BX` | 1 Byte | Suma el contenido de `BX` al registro `AX` (`AX = AX + BX`). |
| **`32h`** | `JNZ dir` | 2 Bytes | Salto condicional: Si `ZF == 0`, establece el `PC = dir`. Si `ZF == 1`, continúa secuencialmente. |
| **`FFh`** | `HLT` | 1 Byte | Detiene la ejecución del procesador (Halt). |

---

## 4. Algoritmo Demostrativo: Multiplicación por Sumas Sucesivas (3 x 4 = 12)

El programa demostrativo cargado en la memoria RAM resuelve la multiplicación de 3 x 4 sumando el valor `04h` un total de 3 veces dentro de un bucle controlado por `BX` y la instrucción `JNZ`.

### 4.1. Código en Ensamblador y Mapeo en RAM

```assembly
00h: MOV AX, 00h    ; Inicializa acumulador AX = 0
02h: MOV BX, 03h    ; Carga el contador de repeticiones BX = 3
04h: ADD AX, 04h    ; [INICIO BUCLE] Suma 4 a AX (AX = AX + 4)
06h: DEC BX         ; Decrementa el contador BX = BX - 1. Si BX == 0, activa ZF = 1
07h: JNZ 04h        ; Si ZF == 0, salta a la dirección 04h
09h: STORE [10h], AX; Guarda el resultado final (0Ch / 12d) en RAM[10h]
0Bh: HLT            ; Detiene el procesador
```

### 4.2. Representación en la RAM (Direcciones `00h` - `0Bh`)

```
10 00 11 03 20 04 23 32 04 13 10 FF
```

---

## 5. Tabla de Traza de Registros y Micro-operaciones

A continuación se presenta la traza paso a paso del ciclo de vida completo de la ejecución del algoritmo:

| Paso | PC | IR | MAR | MDR | AX | BX | ZF | Operación / Estado |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| 0 | 00h | — | — | — | 00h | 00h | 0 | Carga de programa. PC listo en 00h. |
| 1 | 02h | 10h | 00h | 00h | 00h | 00h | 0 | `MOV AX, 00h`: AX se inicializa en 00h. |
| 2 | 04h | 11h | 02h | 03h | 00h | 03h | 0 | `MOV BX, 03h`: BX se inicializa en 03h (contador). |
| 3 | 06h | 20h | 04h | 04h | 04h | 03h | 0 | `ADD AX, 04h` (Iteración 1): AX = 00h + 04h = 04h. |
| 4 | 07h | 23h | 06h | 23h | 04h | 02h | 0 | `DEC BX` (Iteración 1): BX = 03h - 1 = 02h. ZF = 0. |
| 5 | 04h | 32h | 07h | 04h | 04h | 02h | 0 | `JNZ 04h` (Iteración 1): Como ZF == 0, PC salta a 04h. |
| 6 | 06h | 20h | 04h | 04h | 08h | 02h | 0 | `ADD AX, 04h` (Iteración 2): AX = 04h + 04h = 08h. |
| 7 | 07h | 23h | 06h | 23h | 08h | 01h | 0 | `DEC BX` (Iteración 2): BX = 02h - 1 = 01h. ZF = 0. |
| 8 | 04h | 32h | 07h | 04h | 08h | 01h | 0 | `JNZ 04h` (Iteración 2): Como ZF == 0, PC salta a 04h. |
| 9 | 06h | 20h | 04h | 04h | 0Ch | 01h | 0 | `ADD AX, 04h` (Iteración 3): AX = 08h + 04h = 0Ch (12d). |
| 10 | 07h | 23h | 06h | 23h | 0Ch | 00h | 1 | `DEC BX` (Iteración 3): BX = 01h - 1 = 00h. ZF = 1. |
| 11 | 09h | 32h | 07h | 04h | 0Ch | 00h | 1 | `JNZ 04h` (Iteración 3): Como ZF == 1, Fin de bucle. PC = 09h. |
| 12 | 0Bh | 13h | 10h | 0Ch | 0Ch | 00h | 1 | `STORE [10h], AX`: Almacena 0Ch en RAM[10h]. |
| 13 | 0Ch | FFh | 0Bh | FFh | 0Ch | 00h | 1 | `HLT`: Procesador detenido. Estado COMPLETADO. |

---

## 6. Manual de Uso e Interacción

1. **Abrir el archivo de Excel:** Asegurarse de abrir el archivo `Simulador_CPU_VonNeumann.xlsm` habilitando el contenido de macros en la barra superior.
2. **Cargar el programa:** Hacer clic en el botón **Cargar Programa** (o **Reiniciar**). Esto limpiará la RAM, restablecerá todos los registros a sus valores por defecto (00h) y escribirá el programa demostrativo en la dirección `00h`.
3. **Ejecución Automática (RUN):** Hacer clic en el botón **Ejecutar**. El simulador iniciará la secuencia animada, resaltando las celdas de lectura/escritura en la RAM, encendiendo los buses en cian y amarillo, avanzando el puntero del segmento de código ensamblador y actualizando los registros LED.
4. **Verificación de Resultados:** Al finalizar, el panel mostrará el estado **COMPLETADO**, el registro **AX** conservará el valor `0Ch`, la casilla `10h` de la memoria RAM mostrará `0C`, y el Log de Micro-operaciones mantendrá el historial detallado del procesamiento.

---

## 7. Estructura de Archivos del Repositorio

```
├── src/
│   ├── Simulador_CPU_VonNeumann.xlsm   # Libro de Excel habilitado para macros
│   └── Modulo_Simulador.bas            # Código fuente VBA exportado en texto plano
├── docs/
│   └── diagrama_arquitectura.mmd       # Diagrama de arquitectura de bloques en Mermaid
└── README.md                           # Documentación técnica principal del proyecto
```
