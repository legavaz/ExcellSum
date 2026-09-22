from pydantic import BaseModel, Field


class HealthResponse(BaseModel):
    status: str = Field(..., examples=["ok"])


class DryRunResponse(BaseModel):
    log: str = Field(..., examples=["  r 4 [bottom] C: 30 -> =SUM(C2:C3)"])
    dry_run: bool = Field(True, examples=[True])


class ErrorResponse(BaseModel):
    detail: str = Field(..., examples=["Only .xlsx files are supported"])


class AnalyzeResponse(BaseModel):
    sheet: str = Field(..., examples=["TDSheet"])
    sheets: list[str] = Field(..., examples=[["TDSheet"]])
    rows: int = Field(..., examples=[17])
    columns: int = Field(..., examples=[8])
    flat: bool = Field(..., examples=[True])
    grand_total: bool = Field(..., examples=[True])
    total_row: int | None = Field(None, examples=[17])
    value_columns: list[str] = Field(..., examples=[["D", "H"]])
    cells: int = Field(..., examples=[2])
    grouped_rows: int = Field(..., examples=[0])
