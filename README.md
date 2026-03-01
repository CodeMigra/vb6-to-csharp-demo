# vb6 to c#: inventory manager

A migration of a VB6 inventory management form to C# .NET 6. The original was written around 2001 for a small manufacturer running Windows 98/XP. It used ADO for database access (a Jet/Access .mdb file), module-level globals for state, and `On Error GoTo` throughout.

## what the original does

`frmInventory.frm` is a single-form VB6 application. It lets warehouse staff browse, add, edit, and delete inventory items. The database connection and current record state live in module-level `Dim` variables. The save routine builds SQL strings by concatenating user input directly — SQL injection wasn't a concern on a local LAN in 2001. Error handling is `On Error GoTo` with `MsgBox` for every error path.

## what changed

The C# version splits the form logic from the data layer:

- `IInventoryRepository` defines the contract — the form doesn't know or care whether it's talking to SQLite, SQL Server, or a mock
- `SqliteInventoryRepository` replaces the Access MDB with SQLite (no installation required) and uses parameterised queries throughout
- All database calls are `async` — the UI doesn't freeze on slow queries
- `InventoryItem` is a plain data class — no hidden state, no globals
- Error handling uses exceptions with structured try/catch instead of `On Error GoTo`

The form itself (not included here — it's standard WinForms) wires up to `IInventoryRepository` via constructor injection, so you can swap the backend or write tests without touching UI code.

## running it

**VB6** (requires Visual Basic 6 runtime and MDAC):
Open `frmInventory.frm` in the VB6 IDE and press F5. Needs `inventory.mdb` in the same folder.

**C#** (requires .NET 6+):
```bash
dotnet add package System.Data.SQLite
dotnet run
```

## about codemigra

We migrate VB6, COBOL, Fortran, and Delphi to modern stacks. [codemigra.com](https://codemigra.com)
