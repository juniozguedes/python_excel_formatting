from pydantic import BaseModel
import pandas as pd
import datetime


class DocumentResponse(BaseModel):
    data_frame: pd.DataFrame
    start_date: datetime.datetime
    end_date: datetime.datetime

    class Config:
        arbitrary_types_allowed = True
