# jsonl-to-xlsx

Convert JSON Lines table to Excel .xlsx

## Installation

A self-contained single-file binary is [available](../../releases).

Manage installation via mise in **Command Prompt**:

```cmd
winget install --interactive --exact --id Microsoft.Pave
winget install --interactive --exact --id jdx.mise
pave add "%LOCALAPPDATA%/mise/shims"

mise use --global github:thedewi/jsonl-to-xlsx
```

## Usage

```cmd
jsonl-to-xlsx test.xlsx < test.jsonl
```
