import shelve
from CRUD import Create
from CRUD import Read
from CRUD import Update
from CRUD import Delete
from CRUD import Models


class DB:
    def __init__(self) -> None:
        self.db = shelve.open("SAPGUIFramework")
