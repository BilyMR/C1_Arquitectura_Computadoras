# Simulador de CPU de 8 Bits (Arquitectura Von Neumann) en Excel / VBA

**Universidad Católica Boliviana "San Pablo"**
Unidad Académica Regional Santa Cruz
Departamento de Ingenierías y Ciencias Exactas

- **Carrera:** Ingeniería de Software
- **Materia:** Arquitectura de Computadoras (SIS-131)
- **Semestre:** 1/2026
- **Evaluación:** Parcial 1 — Proyecto Práctico y Defensa Oral

---

## 1. Descripción General del Proyecto

Este proyecto consiste en la implementación de un simulador interactivo y visual a nivel de transferencia de registros (RTL) de una CPU de 8 bits basada en la Arquitectura Von Neumann. El sistema está completamente desarrollado en Microsoft Excel utilizando VBA (Visual Basic for Applications) como motor lógico y de animación gráfica.

El simulador modela con rigor técnico:

- El ciclo completo de instrucción (Fetch - Decode - Execute - Store).
- La gestión de memoria RAM de 256 bytes mapeada en formato hexadecimal.
- La interacción entre la Unidad de Control y la ALU (con cálculo de banderas de estado por software).
- La animación en tiempo real del transporte de datos a través de los buses del sistema.

---

## 2. Especificaciones Técnicas del Sistema

### 2.1. Arquitectura de Memoria RAM (256 Bytes: 00h - FFh)

La memoria principal está organizada en una matriz de 16 x 16 celdas de 8 bits, segmentada de la siguiente manera:

| Segmento | Rango | Tamaño | Descripción |
|---|---|---|---|
| Código | 00h a 0Fh | 16 Bytes | Espacio reservado para los opcodes y operandos del programa en ensamblador. |
| Datos | 10h a EFh | 224 Bytes | Espacio para lectura y escritura de variables y resultados del procesamiento. |
| Pila / Stack | F0h a FFh | 16 Bytes | Estructura LIFO utilizada para almacenamiento temporal, con el puntero de pila (SP) inicializado por defecto en FFh. |

### 2.2. Registros Internos del Procesador

- **PC (Program Counter):** Registro de 8 bits que almacena la dirección de memoria de la siguiente instrucción a ejecutar.
- **IR (Instruction Register):** Guarda el Opcode de la instrucción leída durante la fase Fetch.
- **MAR (Memory Address Register):** Mantiene la dirección de RAM sobre la cual se ejecutará un ciclo de lectura o escritura.
- **MDR / MBR (Memory Data Register):** Almacena el dato de 8 bits que viaja desde o hacia la memoria RAM.
- **AX (Acumulador):** Registro multipropósito de 8 bits para resultados aritméticos y lógicos.
- **BX (Registro Base / Contador):** Registro multipropósito de 8 bits, utilizado frecuentemente como contador en bucles.
- **SP (Stack Pointer):** Puntero de 8 bits que apunta al tope de la pila (FFh).

### 2.3. Banderas de Estado sincronizadas por la ALU (AA7)

- **ZF (Zero Flag):** Se establece en 1 si el resultado de la última operación es exactamente 00h; de lo contrario es 0.
- **CF (Carry / Borrow Flag):** Se establece en 1 si ocurre un desbordamiento sin signo (> 255 en sumas) o un préstamo físico (borrow/underflow) al cruzar por debajo de 00h (en restas y decrementos).
- **SF (Sign Flag):** Se establece en 1 si el bit más significativo (Bit 7) del resultado es 1, interpretándose como un valor negativo en complemento a dos.

---

## 3. Conjunto de Instrucciones (ISA)

El procesador implementa una arquitectura optimizada con codificación de longitud variable (1 o 2 bytes) y soporte integral para banderas de estado:

| Opcode (Hex) | Mnemónico / Formato | Tamaño | Descripción Operacional |
|---|---|---|---|
| 10h | `MOV AX, imm` | 2 Bytes | Carga el valor inmediato (imm) en el registro AX. |
| 11h | `MOV BX, imm` | 2 Bytes | Carga el valor inmediato (imm) en el registro BX. |
| 12h | `MOV AX, BX` | 1 Byte | Copia el contenido de BX al registro AX. |
| 13h | `STORE [dir], AX` | 2 Bytes | Almacena el contenido del registro AX en la dirección de RAM [dir]. |
| 14h | `STORE [dir], BX` | 2 Bytes | Almacena el contenido del registro BX en la dirección de RAM [dir]. |
| 15h | `LOAD AX, [dir]` | 2 Bytes | Carga en AX el valor almacenado en la RAM en la dirección [dir]. |
| 16h | `LOAD BX, [dir]` | 2 Bytes | Carga en BX el valor almacenado en la RAM en la dirección [dir]. |
| 20h | `ADD AX, imm` | 2 Bytes | Suma el valor inmediato a AX (AX = AX + imm) y actualiza banderas. |
| 21h | `ADD AX, BX` | 1 Byte | Suma el contenido de BX a AX y actualiza banderas. |
| 22h | `INC AX` | 1 Byte | Incrementa AX en 1 y actualiza banderas. |
| 23h | `DEC BX` | 1 Byte | Decrementa BX en 1; actualiza ZF, CF por préstamo y SF. |
| 24h | `DEC AX` | 1 Byte | Decrementa AX en 1; actualiza ZF, CF por préstamo y SF. |
| 25h | `SUB AX, imm` | 2 Bytes | Resta un valor inmediato a AX y actualiza banderas de préstamo. |
| 26h | `SUB AX, BX` | 1 Byte | Resta BX a AX y actualiza banderas de préstamo. |
| 27h | `CMP AX, imm` | 2 Bytes | Compara AX con un valor inmediato actualizando banderas sin alterar AX. |
| 28h | `CMP AX, BX` | 1 Byte | Compara AX con BX actualizando banderas. |
| 30h | `JMP dir` | 2 Bytes | Salto incondicional: establece el PC = dir. |
| 31h | `JZ dir` | 2 Bytes | Salto condicional: si ZF == 1, establece el PC = dir. |
| 32h | `JNZ dir` | 2 Bytes | Salto condicional: si ZF == 0, establece el PC = dir. |
| FFh | `HLT` | 1 Byte | Detiene la ejecución del procesador (Halt). |

---

## 4. Programas Demostrativos Incluidos

El simulador cuenta con rutinas integradas para validar distintos comportamientos de hardware y algoritmos de nivel académico:

1. **Factorial de 4 (4! = 24 = 18h):** Realiza el cálculo matemático mediante fases de sumas sucesivas estructuradas, almacenando el resultado final directamente en la celda de memoria [20h].
2. **Prueba de Underflow:** Ejecuta un conteo descendente a través de los límites de los 8 bits en complemento a dos, evidenciando el encadenamiento síncrono del Carry Flag durante el préstamo y la detención limpia tras 256 iteraciones.
3. **Acumulación por Suma Sucesiva:** Simula operaciones aritméticas incrementales utilizando la instrucción ADD y saltos condicionales.

---

## 5. Manual de Uso e Interacción

1. **Abrir el archivo de Excel:** Asegurarse de abrir el libro de trabajo `.xlsm` habilitando el contenido de macros en la barra de seguridad superior.
2. **Cargar un programa:** Hacer clic en el botón correspondiente a la rutina de carga (ej. *Cargar Factorial* o *Cargar Prueba*). Esto limpiará la RAM, restablecerá los registros y escribirá el código ensamblador en las direcciones iniciales.
3. **Ejecución Automática (RUN) o Paso a Paso (STEP):** Iniciar la ejecución mediante los controles de la interfaz. El simulador animará los buses de direcciones (cian) y datos (amarillo), resaltará las celdas activas en la matriz de la RAM y actualizará los registros LED y el panel de Log cronológico.
4. **Verificación de Resultados:** Al finalizar con la instrucción `HLT`, el panel indicará **COMPLETADO**, los registros reflejarán el estado final de la computación y las banderas (AA7) mostrarán los indicadores lógicos de la ALU.

---

## 6. Estructura de Archivos del Repositorio

```text
├── src/
│   ├── Simulador_CPU_VonNeumann.xlsm   # Libro de Excel habilitado para macros
│   └── Modulos_VBA/                    # Código fuente exportado (Módulos 1, 2 y 3)
├── docs/
│   └── diagrama_arquitectura.mmd       # Diagrama de arquitectura de bloques en Mermaid
└── README.md                           # Documentación técnica principal del proyecto
```
