# Excel Helpers

A collection of documented and unit tested named [`LAMBDA`](https://support.microsoft.com/en-gb/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) functions for Excel.

Reusable functions help with:

- Consistency in the way data is calculated and formatted
- Bugs where you perhaps write a formula slightly different in two places
- A greater level of trust in calculations, particularly in large documents
- If you need to change a calculation, you just change it in a single place

Each [`LAMBDA`](https://support.microsoft.com/en-gb/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) function is documented in Markdown, and contains a downloadable Excel spreadsheet which contains the unit tests, so you can trust what it's doing and its expected outputs.

## Contributing

If you want to contribute to this project, here are the rules:

Every function must:

- Be documented as a standalone Markdown file
- Have robust unit tests in a downloadable file using the template
- Follow the defined structure conventions
- Follow the defined naming conventions
- Follow the correct error handling conventions

## Function structure conventions

Functions in this project should always follow these structure conventions. It does make the code more long-winded, but it makes them all feel familiar, easier to read and easier to debug.

Each function should be structured using the following conventions:

1. Functions follow the readable cascade pattern
2. Arguments are defined using naming conventions
3. Arguments follow the argument handling pattern
4. Errors are defined using naming conventions
5. Errors follow the error handling pattern
6. Results are defined using naming conventions
7. Outputs are defined using naming conventions
8. Outputs follow the output handling pattern

## Patterns

Patterns define the underlying logic behind all functions in this project. They ensure every function is readable, predictable, consistent, and debuggable. Every Excel Helper function in this project must follow these patterns.

### Readable cascade pattern

The readable cascade pattern is the foundation of every function in this project. It ensures that logic flows in a clear, top-to-bottom sequence and that every step can be understood in isolation.

#### Principles

- Each LET block defines one new variable per line
- Variables flow sequentially, with later variables depending on earlier ones
- No deeply nested logic inside a single assignment
- "Verbosity" over "cleverness", clarity is always better
- Cascade ordering:
  1. Assign arguments
  2. Validate arguments
  3. Perform core logic using validated arguments
  4. Combine errors after all validation is complete
  5. Compute the final `_result`
  6. Set the final `_output`
  7. Return `_output` as the final step

#### Purpose

- Makes functions easy to read for everyone, especially those less experienced
- Helps to prevent hidden logic
- Makes debugging easier, inspect line-by-line variable-by-variable
- Standardises structure across the entire library

Use this example pattern to ensure you write consistent and predictable functions.

```js
=LAMBDA([arg1],
  LET(
    _arg1, ...,     // Assign arguments
    _ERR_arg1, ..., // Validate arguments
    _core, ...,     // Define core logic
    _ERR_msg, ...,  // Capture errors
    _result, ...,   // Compute result
    _output, ...,   // Calculate the final output
    _output         // Return the final output
  )
)
```

### Argument handling pattern

Arguments must be validated in a consistent and predictable way. The argument handling pattern ensures that:

- All arguments behave the same way
- All errors are caught and defined
- No native Excel errors leak out and confuse the user
- Contributors don’t need to reinvent validation logic

#### Rules

1. All arguments must be wrapped in square brackets when defined in the [`LAMBDA`](https://support.microsoft.com/en-gb/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) signature. For example:
    ```js
    =LAMBDA([number],...
    ```

2. Inside the LET cascade, arguments must be assigned to internal variables with the same name, prefixed with `_`. For example:
    ```js
    _number, number,
    ```

3. Required arguments must evaluate to `NA()` when omitted, so they can be validated. For example:
    ```js
    _number, IF(ISOMITTED(number), NA(), number)
    ```

4. Required arguments which are invalid must be assigned clear error messages. For example:  
    ```js
    _number, IF(ISOMITTED(number), NA(), number)
    _ERR_number, IFS(
      ISNA(_number), "[number] argument is omitted",
      NOT(ISNUMBER(_number)), "[number] argument is not a number",
      TRUE, ""
    ),
    ```

4. Optional arguments must evaluate to a safe default when omitted, depending on the type. For example:
    ```js
    _text, IF(ISOMITTED(text), "", text),
    _boolean, IF(ISOMITTED(boolean), FALSE, boolean),
    _number, IF(ISOMITTED(number), 0, number),
    ```

6. Optional arguments which are invalid must *not* be validated for omission, but should still be evaluated for anything else, like type. For example:  
    ```js
    _prefix, IF(ISOMITTED(prefix), "", prefix),
    _ERR_prefix, IFS(
      NOT(ISTEXT(_prefix)), "[prefix] argument is not text",
      TRUE, ""
    ),
    ```

7. After validation, never reference the raw argument again. Always refer to the defined variable. For example:
    ```js
    =LAMBDA([number],
      _number, IF(ISOMITTED(number), NA(), number),
      // Always use _number from now on, not number
    )
    ```

#### Purpose

- Ensures inputs are always predictable
- Prevents unexpected raw user input reaching core logic
- Allows error handling to work uniformly across the library

### Error handling pattern

Error handling in this project is strict, deliberate, and standardised.

It ensures that:
- All errors are returned as useful, readable text
- No errors propagate inside Excel’s calculation engine
- Complex functions can aggregate multiple error messages
- The final `_output` is more predictable

#### Principles

- Each validation block produces an error message
- A single validation block is defined as `_ERR_msg`
- Multiple error blocks are defined as `_ERR_variableName`, then combined into a single `_ERR_msg` variable
- The combination should always follow this sequence:
    ```js
    _ERR_msg, LET(
      _errors, VSTACK(_ERR_var1, _ERR_var2, ...),
      _cleaned, FILTER(_errors, _errors<>"", ""),
      _line1, "ERROR:",
      _prefix, " • ",
      _joined, TEXTJOIN(CHAR(10) & _prefix, TRUE, _cleaned),
      _combined, _line1 & CHAR(10) & _prefix & _joined,
      _combined
    )
    ```
- `_output` variable must prioritise `_ERR_msg` over `_result`
- Native Excel errors should never escape the function
- Only text-based errors should be returned

#### Purpose

- Provides consistent, human-readable error messages
- Works well with unit testing
- Avoids duplicated logic and inconsistent validation styles
- Makes functions safe to compose together

## Naming conventions

### Functions

We define them once using the [`LAMBDA`](https://support.microsoft.com/en-gb/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) keyword. Then save them as named functions with a unique name to describes their purpose clearly.

**Rules**
- **Text**: [PascalCase](https://www.theserverside.com/definition/Pascal-case)
- **Prefix**: Not allowed
- **Postfix**: Not allowed
- **Reasoning**: PascalCase helps differentiate custom functions from native Excel functions, which use Uppercase. Functions should not use a prefix or postfix with symbols or numbers, they should have unique names which describe their purpose.

**Definition example**:   
`FunctionName`

**Conceptual examples**:
```js
=CalculatePercentage(10, 100)

=Multiply(5, 10)
```

### Arguments

Arguments are the values we pass into functions. They should have a name which describes their expectation clearly. As per the arguments handling pattern, they should always be wrapped in square brackets when defined as part of a [`LAMBDA`](https://support.microsoft.com/en-gb/office/lambda-function-bd212d27-1cd1-4321-a34a-ccbf254b8b67) function.

**Rules**
- **Text**: [camelCase](https://www.techtarget.com/whatis/definition/CamelCase)
- **Prefix**: Not allowed
- **Postfix**: Numbers allowed
- **Reasoning**: camelCase helps differentiate arguments from functions, which use Uppercase or PascalCase. Where you need to define multiple arguments which are similar, you can postfix them with numbers.

**Definition example**:   
`argName1`

**Conceptual examples**:
```js
// LAMBDA replaced with FunctionName for example

=Divide([smallNumber], [bigNumber])

=Multiply([number1], [number2])
```

### Standard variables

Standard variables are anything which is not an error, result or the final output from the function. They are used to store temporary values and pass them down through the cascade, making the code easier to read.

**Rules**
- **Text**: [camelCase](https://www.techtarget.com/whatis/definition/CamelCase)
- **Prefix**: `_` (Underscore)
- **Postfix**: Numbers allowed
- **Reasoning**: camelCase helps differentiate variables from functions, which use Uppercase or PascalCase. An underscore prefix helps differentiate from arguments, which also use camelCase.

**Definition example**:   
`_variableName1`

**Conceptual example**:
```js
=LET(
  // Sets _upperText to "SOME UPPERCASE TEXT"
  _upperText, "SOME UPPERCASE TEXT",

  // Sets _lowerText to "some uppercase text"
  _lowerText, LOWER(_upperText),

  // Sets _result to "some lowercase text"
  _result, SUBSTITUTE(_lowerText, "uppercase", "lowercase"),

  // Sets _output to _result to follow convention
  _output, _result,
  
  // Returns the final _output and displays it in the cell
  _output
)
```

### Result variable

The `_result` variable is reserved for the final computed value of the function. This is not *necessarily* what gets outputted as the final return value, but it's the very last step in the core logic.

**Rules**
- **Text**: `result` (as written)
- **Prefix**: `_` (underscore)
- **Postfix**: Not allowed
- **Reasoning**: The `_result` variable is used to store the final computed value from the core logic. By always assigning it to a static named variable, it makes it easy to spot the end of the core logic, which is useful for debugging. Numbers are not allowed as a postfix as you should only collect the final computed result once.

**Definition example**:   
`_result`

**Conceptual example**:
```js
LET(
  // Sets _number1 to 5
  _number1, 5,

  // Sets _number2 to 10
  _number2, 10,

  // Sets _result to 50, which is (5 * 10)
  _result, (_number1 * _number2),

  // Sets _output to _result to follow convention
  _output, _result,

  // Returns the final _output and displays it in the cell
  _output
)
```

### Output variable

The `_output` variable is reserved for the final return value of the function. The `_output` can be equal to the `_result`, or in the case of errors, it can be equal to the `_ERR_msg`.

**Rules**
- **Text**: `output` (as written)
- **Prefix**: `_` (underscore)
- **Postfix**: Not allowed
- **Reasoning**: The `_output` variable is used to store the final return value from the function. This is the value the function actually returns and is shown in the cell. By always assigning it to a static named variable, it makes it easy to spot the expected point of return in every function. Numbers are not allowed as a postfix as you should only collect the final output once at the very end of the function.

**Definition example**:   
`_output`

**Conceptual example**:
```js
LET(
  // Sets _number1 to 5
  _number1, 5,

  // Sets _number2 to 10
  _number2, 10,

  // Sets _result to 50, which is (5 * 10)
  _result, (_number1 * _number2),

  // Sets _output to _result to follow convention
  _output, _result,

  // Returns the final _output and displays it in the cell
  _output
)
```

### Error message variable

The error message variable is reserved for articulating errors based on validation checks. The variable name `_ERR_msg` is reserved for the final combined error message.

For complex functions that collect multiple error messages at different stages, use the naming convention `_ERR_variableName` for each logic check, and then combine all the error messages at the end under one `_ERR_msg` variable.

Error message variables should always be defined close to their logic checks in order to keep the code readable.

**Rules**
- **Text**: `_msg` or `_variableName` (inherited)
- **Prefix**: `_ERR`
- **Postfix**: Numbers allowed
- **Reasoning**: Error variables are defined using the prefix `_ERR` followed by the variable name they refer to. This helps differentiate them from standard variables and also makes it clear which variable they relate to. If you need to define more than one error for the same variable you can use numbers as a postfix.

**Definition example**:   
`_ERR_variableName`

**Simple conceptual example**:
```js
// LAMBDA replaced with FunctionName for example
=CheckIfNumber([input],
  // Starts the LET cascade
  LET(
    // Sets _input value if valid, or as error if not
    _input, IF(ISOMITTED(input), NA(), input),

    // Sets _ERR_msg to the outcome of the IF logic
    _ERR_msg, IF(
      // If _input is omitted, return this error message
      ISNA(_input), "[input] argument is omitted",
      // If it's not, return empty string
      ""
    ),

    // Sets _result to TRUE or FALSE depending on if _input is a number
    _result, ISNUMBER(_input),

    // _output is set to the outcome of the IF logic
    _output, IF(
      // If _ERR_msg contains more than 0 characters
      LEN(_ERR_msg) > 0,
      // Then return _ERR_msg (our error message)
      _ERR_msg,
      // Otherwise, return _result
      _result
    ),

  // Returns the final _output and displays it in the cell
    _output
  )
)
```

**Complex conceptual example**:
```js
// LAMBDA replaced with FunctionName for example
=Multiply([number1],[number2],
    // Starts the LET cascade
    LET(
    // Sets _number1 value if valid, or as error if not
    _number1, IF(ISOMITTED(number1), NA(), number1),

    // Sets _ERR_number1 to the outcome of the IF logic
    _ERR_number1, IFS(
      // If _number1 is omitted, return this error message
      ISNA(_number1), "[number1] argument is omitted",
      // If _number1 is not a number, return this error message
      NOT(ISNUMBER(_number1)), "[number1] argument is not a number",
      // If _number1 is valid, return empty string
      TRUE, ""
    ),

    // Sets _number2 value if valid, or as error if not
    _number2, IF(ISOMITTED(number2), NA(), number2),
    _ERR_number2, IFS(
      // If _number2 is omitted, return this error message
      ISNA(_number2), "[number2] argument is omitted",
      // If _number2 is not a number, return this error message
      NOT(ISNUMBER(_number2)), "[number2] argument is not a number",
      // If _number2 is valid, return empty string
      TRUE, ""
    ),

    // Use a second LET cascade to define and combine _ERR_msg
    _ERR_msg, LET(
      // Sets _errors to be a combined array of errors
      _errors, VSTACK(_ERR_number1, _ERR_number2),
      // Removes any empty values from the array
      _cleaned, FILTER(_errors, _errors<>"", ""),
      // Combines the error messages with a "–" separator
      _joined, TEXTJOIN(" – ", TRUE, _cleaned),
      // Returns _joined and sets it as the value of _ERR_msg
      _joined
    ),

    // Sets _result to the value of _number1 * _number2
    _result, (_number1 * _number2),

    // Sets the value of _output to the outcome of the IF logic
    _output, IF(
      // If _ERR_msg contains more than 0 characters
      LEN(_ERR_msg) > 0,
      // Sets _output to be the _ERR_msg
      _ERR_msg,
      // Otherwise, set _output to be the _result
      _result
    ),
    
  // Returns the final _output and displays it in the cell
    _output
  )
)
```