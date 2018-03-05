# Netplate
Text template/mail merge written in C#

# Usage

## Basic mail merge
The basic mail merge uses the format `[key]` where `key` is the key in the `data` dictionary.
The entire text `[key]` is replaced by the associated `value.ToString()` of `data`.

Simple conditional content can be added before and after `[key]` using the format `[Content~]` for before and `[~Content]` for after. The brackets `[]` and tilde `~` are removed regardless, and the content is only shown if the value is none null/empty.

Conditional content is optional and can be added before and/or after.

Conditional content may contain additional merge fields and matching conditional content.

### Examples
`Dear [Surname],` where `Surname` = `Mr Smith` results in `Dear Mr Smith,`

`[Before ~][My Value][~ After]` where `My Value` = `Middle` results in `Before Middle After`

`Normal text [Value A][~ [Value B]]` where `Value A` = `1` and `Value B` = `2` results in `Normal text 1 2`

`Normal text [Value A][~ [Value B]]` where `Value A` = `<null/empty>` and `Value B` = `2` results in `Normal text`

## `if` mail merge
The `if` mail merge uses the format `[if <operand> <operator> <operand>]`.

Each operand is replaced with either a mail merge key `[key]`, a literal integer `12`, a literal decimal `12.52` - or a literal string `"My string"`.

The accepted operators are:
- `==` Equal to
- `!=` Not equal to
- `>` Greater than
- `<` Less than
- `>=` Greater than or equal to
- `<=` Less than or equal to

Conditional content is formatted the same way as with basic mail merge: `[Before~][if <operand> <operator> <operand>][~After]`