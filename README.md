# Kwinto Analytics - XLL Add-In

## User Manual

This project is in progress. Some functions below are implemented partly or not at all.

### Create JSON

```
=kwJson([id], key1, val1, ... key6, val6) -> "json:<id>" | "temp:<uid>"
```

`kwJson` function creates a JSON object `{key1: val1, ... key6: val6}` in Excel memory under
`"json:<id>"` name, which you can use in other `kw*` functions later, where JSON argument is
required. This object is stored _permanently_, so that you can use it many times.

When `id` is missing, `kwJson` creates JSON under `temp:<uid>` name. This object is stored
_temporary_, so that after the first use it will be removed. This helps to manage memory,
avoid name collisions, and correctly build Excel dependency tree.

Example:

```
// Create JSON
=kwJson("config", "port", B2)
=kwJson("data", A2, B2:B4)

// Copy JSON from temporary to permanent
=kwJson("config",, "temp:12")

// Merge two JSON objects (note missing <key1> and <key2>)
=kwJson("config", ,"json:config", ,kwJson("", "host", "gituliar.net"))
```

### Send JSON-RPC

```
=kwRpc(method, id, params) -> "json:<id>"
```

`kwRpc` invokes a remote function, waits for it to complete, and saves its result under
`"json:<id>"` name, which you can later use in other `kw*` functions. On error a plain text message
is returned instead.

### Show JSON

```
=kwShow(id) -> "<json>"
```

`kwShow` return JSON object serialized to a string.

Example:

```
// Show temporary JSON
=kwShow(kwJson(, A2, B2:B4))
```

### Get Value from JSON

```
=kwValue(id, key) -> "<value>"
```

### Config

```
=kwConfig() -> "<json>"
=kwConfig(id) -> "<json>"
```

## Useful Links

- Excel Software Development Kit

  https://learn.microsoft.com/en-us/office/client-developer/excel/welcome-to-the-excel-software-development-kit

- JSON-RPC Specification

  https://www.jsonrpc.org/specification

  https://www.jsonrpc.org/historical/
