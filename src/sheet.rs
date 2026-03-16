use physis::{
    excel::{Field, Row},
    resource::ResourceResolver,
};

#[derive(thiserror::Error, Debug)]
pub enum SheetError {
    #[error("Failed to read EXH header")]
    HeaderReadError,

    #[error("Failed to read EXD data")]
    DataReadError,

    #[error("Row not found: {0}")]
    RowNotFound(u32),

    #[error("Subrow not found: {0}, {1}")]
    SubrowNotFound(u32, u16),
}

#[derive(thiserror::Error, Debug)]
pub enum ColumnReadError {
    #[error("Unexpected column type: got {0:?}, expected: {1}")]
    UnexpectedType(Field, &'static str),
}

pub trait FromPhysisRow {
    fn from_single_row(row: &Row) -> Self;
}

pub trait FromExcelRow<R: FromPhysisRow>: Sized {
    fn from_row(row: R) -> Self;
}

pub struct RowRef<R> {
    pub row_id: u32,
    pub subrow_id: u16,
    pub row: R,
}

impl<R> RowRef<R> {
    pub fn row(row_id: u32, row: R) -> Self {
        Self {
            row_id,
            subrow_id: 0,
            row,
        }
    }

    pub fn subrow(row_id: u32, subrow_id: u16, row: R) -> Self {
        Self {
            row_id,
            subrow_id,
            row,
        }
    }
}

pub trait Sheet: Sized {
    type Row: 'static + FromPhysisRow;

    const NAME: &'static str;

    fn from_sheet(sheet: physis::excel::Sheet) -> Self;

    fn get_sheet(&self) -> &physis::excel::Sheet;

    fn read_from(
        resolver: &mut ResourceResolver,
        language: physis::Language,
    ) -> Result<Self, SheetError> {
        let exh = resolver
            .read_excel_sheet_header(Self::NAME)
            .map_err(|_| SheetError::HeaderReadError)?;

        let sheet = resolver
            .read_excel_sheet(&exh, Self::NAME, language)
            .map_err(|_| SheetError::DataReadError)?;

        Ok(Self::from_sheet(sheet))
    }

    fn page_index_for_row(&self, row_id: u32) -> Option<usize> {
        self.get_sheet()
            .exh
            .pages
            .iter()
            .position(|p| row_id >= p.start_id && row_id < p.start_id + p.row_count)
    }

    fn get_row(&self, row_id: u32) -> Result<Self::Row, SheetError> {
        match self.get_sheet().row(row_id) {
            Some(row) => Ok(Self::Row::from_single_row(row)),
            _ => Err(SheetError::RowNotFound(row_id)),
        }
    }

    fn get_subrow(&self, row_id: u32, subrow_id: u16) -> Result<Self::Row, SheetError> {
        let row = self
            .get_sheet()
            .subrow(row_id, subrow_id)
            .ok_or(SheetError::SubrowNotFound(row_id, subrow_id))?;

        Ok(Self::Row::from_single_row(row))
    }

    fn row_count(&self) -> u32 {
        self.get_sheet().exh.header.row_count
    }

    fn iter(&self) -> Box<dyn Iterator<Item = RowRef<Self::Row>> + '_> {
        let sheet = self.get_sheet();

        match sheet.exh.header.row_kind {
            physis::exh::SheetRowKind::SingleRow => Box::new(sheet.pages.iter().flat_map(|page| {
                page.entries.iter().filter_map(|entry| {
                    self.get_row(entry.id)
                        .ok()
                        .map(|row| RowRef::row(entry.id, row))
                })
            })),

            physis::exh::SheetRowKind::SubRows => Box::new(sheet.pages.iter().flat_map(|page| {
                page.entries.iter().flat_map(|entry| {
                    entry.subrows.iter().map(|(subrow_id, data)| {
                        RowRef::subrow(entry.id, *subrow_id, Self::Row::from_single_row(data))
                    })
                })
            })),
        }
    }

    fn iter_as<T>(&self) -> impl Iterator<Item = RowRef<T>> + '_
    where
        T: FromExcelRow<Self::Row>,
    {
        self.iter().map(|rowref| RowRef {
            row_id: rowref.row_id,
            subrow_id: rowref.subrow_id,
            row: T::from_row(rowref.row),
        })
    }

    fn get_row_as<T>(&self, row_id: u32) -> Result<T, SheetError>
    where
        T: FromExcelRow<Self::Row>,
    {
        self.get_row(row_id).map(T::from_row)
    }

    fn get_subrow_as<T>(&self, row_id: u32, subrow_id: u16) -> Result<T, SheetError>
    where
        T: FromExcelRow<Self::Row>,
    {
        self.get_subrow(row_id, subrow_id).map(T::from_row)
    }
}
