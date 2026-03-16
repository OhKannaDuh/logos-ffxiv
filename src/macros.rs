#[macro_export]
macro_rules! define_sheet {
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

// FIELDS
#[macro_export]
macro_rules! string_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<String, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::String(s) => Ok(s.clone()),
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
macro_rules! u8_field {
    ($fn_name:ident, $idx:expr) => {
        pub fn $fn_name(&self) -> Result<u8, $crate::ColumnReadError> {
            match &self[$idx] {
                physis::excel::Field::UInt8(s) => Ok(s.clone()),
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
                physis::excel::Field::UInt16(s) => Ok(s.clone()),
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
                physis::excel::Field::UInt32(s) => Ok(s.clone()),
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
                physis::excel::Field::UInt64(s) => Ok(s.clone()),
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
                physis::excel::Field::Int8(s) => Ok(s.clone()),
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
                physis::excel::Field::Int16(s) => Ok(s.clone()),
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
                physis::excel::Field::Int32(s) => Ok(s.clone()),
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
                physis::excel::Field::Int64(s) => Ok(s.clone()),
                other => Err($crate::ColumnReadError::UnexpectedType(
                    other.clone(),
                    "i64",
                )),
            }
        }
    };
}

pub use i64_field;

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
