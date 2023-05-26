# spo tenant commandset set

Updates a ListView Command Set that is installed tenant wide.

## Usage

```sh
spo tenant commandset set [options]
```

## Options

`-i, --id <id>`
: The id of the ListView Command Set

`-t, --newTitle [newTitle]`
: The updated title of the ListView Command Set

`-l, --listType [listType]`
: The list or library type to register the ListView Command Set on. Allowed values `List` or `Library`.

`-i, --clientSideComponentId  [clientSideComponentId]`
: The Client Side Component Id (GUID) of the ListView Command Set.

`-p, --clientSideComponentProperties  [clientSideComponentProperties]`
: The Client Side Component properties of the ListView Command Set.

`-w, --webTemplate [webTemplate]`
: Optionally add a web template (e.g. STS#3, SITEPAGEPUBLISHING#0, etc) as a filter for what kind of sites the ListView Command Set is registered on.

`--location [location]`
: The location of the ListView Command Set. Allowed values `ContextMenu`, `CommandBar` or `Both`. Defaults to `CommandBar`.

--8<-- "docs/cmd/_global.md"

## Remarks

!!! warning "Escaping JSON in PowerShell"
    When using the `--clientSideComponentProperties` option it's possible to enter a JSON string. In PowerShell 5 to 7.2 [specific escaping rules](./../../../user-guide/using-cli.md#escaping-double-quotes-in-powershell) apply due to an issue. Remember that you can also use [file tokens](./../../../user-guide/using-cli.md#passing-complex-content-into-cli-options) instead.

## Examples

Updates the title of a ListView Command Set that's deployed tenant wide.

```sh
m365 spo tenant commandset set --id 4  --newTitle "Some customizer"
```

Updates the properties of a ListView Command Set.

```sh
m365 spo tenant commandset  set --id 3  --clientSideComponentProperties '{ "someProperty": "Some value" }'
```

## Response

The command won't return a response on success.