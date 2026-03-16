# Logos

Logos is a Rust crate for working with Final Fantasy XIV’s EXD data files, built on top of the Physis backend. The project draws inspiration from [Icarus](https://github.com/redstrate/Icarus) while offering a more opinionated API aligned with my own design preferences. If Logos doesn’t suit your needs, I recommend taking a look at Icarus as an alternative.

# Reading a sheet

```rs
let mut resolver = ResourceResolver::new();

let territories =
    TerritoryTypeSheet::read_from(&mut resolver, physis::Language::None);
// Result<TerritoryTypeSheet, SheetError>
```

# Accessing data

```rs
// Get a row from a sheet without subrows
territories.get_row(1190); // Result<TerritoryTypeRow, SheetError>

// Get a subrow from a sheet that supports subrows
map_markers.get_subrow(3, 12); // Result<MapMarkerRow, SheetError>
```

# Iterating over a sheet

```rs
for RowRef { row_id, row, .. } in territories.iter() {
    // ...
}

for RowRef { row_id, subrow_id, row } in map_markers.iter() {
    // ...
}
```

# Typed data views

You can define your own strongly typed wrappers by implementing the FromExcelRow trait:

```rs
pub trait FromExcelRow<R: FromPhysisRow>: Sized {
    fn from_row(row: R) -> Self;
}
```

This enables accessing sheet data through your own types. Logos provides three typed‑view helpers:

```rs
    iter_as<T>()

    get_row_as<T>()

    get_subrow_as<T>()
```

These behave like their untyped counterparts but automatically convert rows into your wrapper type.

## Example

```rs
pub struct Territory(TerritoryTypeRow);

// A convenience macro exists for this:
// from_excel_row_newtype!(Territory, TerritoryTypeRow);
impl FromExcelRow<TerritoryTypeRow> for Territory {
    fn from_row(row: TerritoryTypeRow) -> Self {
        Self(row)
    }
}

let territory = territories.get_row_as::<Territory>(1190);
let marker = map_markers.get_subrow_as::<SomeMapMarkerWrapper>(3, 12);

for RowRef { row_id, row: territory, .. } in territories.iter_as::<Territory>() {
    // `territory` is of type `Territory`
}
```
