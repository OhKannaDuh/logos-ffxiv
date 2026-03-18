macro_rules! sheet_common {
    ($sheet_ty:ident, $row_ty:ty, $name:literal) => {
        #[derive(Debug, Clone)]
        pub struct $sheet_ty {
            sheet: physis::excel::Sheet,
        }

        impl Sheet for $sheet_ty {
            type Row = $row_ty;
            const NAME: &'static str = $name;

            fn from_sheet(sheet: physis::excel::Sheet) -> Self {
                Self { sheet }
            }

            fn get_sheet(&self) -> &physis::excel::Sheet {
                &self.sheet
            }
        }
    };
}

pub(crate) use sheet_common;

#[macro_export]
macro_rules! define_sheet {
    ($sheet_ty:ident, $row_ty:ty, $name:literal, language_support) => {
        $crate::macros::sheet_common!($sheet_ty, $row_ty, $name);

        impl $sheet_ty {
            pub fn read_from(
                resolver: &mut physis::resource::ResourceResolver,
                language: physis::Language,
            ) -> Result<Self, $crate::SheetError> {
                let exh = resolver
                    .read_excel_sheet_header(Self::NAME)
                    .map_err(|_| $crate::SheetError::HeaderReadError)?;

                let sheet = resolver
                    .read_excel_sheet(&exh, Self::NAME, language)
                    .map_err(|_| $crate::SheetError::DataReadError)?;

                Ok(Self::from_sheet(sheet))
            }
        }
    };

    ($sheet_ty:ident, $row_ty:ty, $name:literal, no_language_support) => {
        $crate::macros::sheet_common!($sheet_ty, $row_ty, $name);

        impl $sheet_ty {
            pub fn read_from(
                resolver: &mut physis::resource::ResourceResolver,
            ) -> Result<Self, $crate::SheetError> {
                let exh = resolver
                    .read_excel_sheet_header(Self::NAME)
                    .map_err(|_| $crate::SheetError::HeaderReadError)?;

                let sheet = resolver
                    .read_excel_sheet(&exh, Self::NAME, physis::Language::None)
                    .map_err(|_| $crate::SheetError::DataReadError)?;

                Ok(Self::from_sheet(sheet))
            }
        }
    };
}

pub use define_sheet;

#[macro_export]
macro_rules! define_row {
    ($row_ty:ident) => {
        #[derive(Debug, Clone)]
        pub struct $row_ty(physis::excel::Row);
        impl core::ops::Deref for $row_ty {
            type Target = [physis::excel::Field];
            fn deref(&self) -> &Self::Target {
                &self.0.columns
            }
        }

        impl $crate::FromPhysisRow for $row_ty {
            fn from_single_row(row: &physis::excel::Row) -> Self {
                Self(row.clone())
            }
        }
    };
}

pub use define_row;

#[macro_export]
macro_rules! define_subrow {
    ($row_ty:ident, $field_count:expr) => {
        #[derive(Debug, Copy, Clone)]
        pub struct $row_ty<'a>(&'a [physis::excel::Field]);
        impl<'a> core::ops::Deref for $row_ty<'a> {
            type Target = [physis::excel::Field];
            fn deref(&self) -> &Self::Target {
                self.0
            }
        }

        impl<'a> $row_ty<'a> {
            pub const FIELD_COUNT: usize = $field_count;
        }
    };
}

pub use define_subrow;

// FIELDS
#[macro_export]
macro_rules! array_field {
    ($fn_name:ident, $idx:expr, $len:expr, $inner:ident) => {
        pub fn $fn_name(&self) -> impl Iterator<Item = $inner<'_>> {
            (0..$len).map(|i| {
                let start = $idx + i * $inner::FIELD_COUNT;
                let end = start + $inner::FIELD_COUNT;
                $inner(&self[start..end])
            })
        }
    };
}

pub use array_field;

#[macro_export]
macro_rules! string_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<&str, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::String(s) => Ok(s.as_str()),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "String",
                )),
            }
        }
    };
}

pub use string_field;

#[macro_export]
macro_rules! f32_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<f32, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::Float32(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "f32",
                )),
            }
        }
    };
}

pub use f32_field;

#[macro_export]
macro_rules! u8_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<u8, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::UInt8(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(other.clone(), "u8")),
            }
        }
    };
}

pub use u8_field;

#[macro_export]
macro_rules! u16_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<u16, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::UInt16(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "u16",
                )),
            }
        }
    };
}

pub use u16_field;

#[macro_export]
macro_rules! u32_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<u32, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::UInt32(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "u32",
                )),
            }
        }
    };
}

pub use u32_field;

#[macro_export]
macro_rules! u64_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<u64, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::UInt64(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "u64",
                )),
            }
        }
    };
}

pub use u64_field;

#[macro_export]
macro_rules! i8_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<i8, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::Int8(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(other.clone(), "i8")),
            }
        }
    };
}

pub use i8_field;

#[macro_export]
macro_rules! i16_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<i16, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::Int16(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "i16",
                )),
            }
        }
    };
}

pub use i16_field;

#[macro_export]
macro_rules! i32_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<i32, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::Int32(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "i32",
                )),
            }
        }
    };
}

pub use i32_field;

#[macro_export]
macro_rules! i64_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<i64, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::Int64(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "i64",
                )),
            }
        }
    };
}

pub use i64_field;

#[macro_export]
macro_rules! bool_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<bool, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::Bool(v) => Ok(*v),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "bool",
                )),
            }
        }
    };
}

pub use bool_field;

// HELPERS
#[macro_export]
macro_rules! from_excel_row_newtype {
    ($wrapper:ty, $inner:ty) => {
        impl $crate::FromExcelRow<$inner> for $wrapper {
            fn from_row(row: $inner) -> Self {
                Self(row)
            }
        }
    };
}

pub use from_excel_row_newtype;
