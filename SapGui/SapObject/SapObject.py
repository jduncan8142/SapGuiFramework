import json
import time
import win32com.client
from typing import Any, Optional
from thefuzz import fuzz, process


class SapObject:
    def __init__(self, session: win32com.client.CDispatch, id: Optional[str] = None) -> None:
        self.session: win32com.client.CDispatch = session
        self.tree: dict = json.loads(session.getObjectTree(session.id))
        self.id_list: list = []
        self.enumerate_object_tree(tree=self.tree)
        self.id: str | None = id if id is not None else None
        self.element: win32com.client.CDispatch | None 
        self.is_container: bool
        self.name: str
        self.parent: win32com.client.CDispatch | None
        self.type: str
        self.type_as_number: int
        if self.id is not None:
            self.get_element()        

    def enumerate_object_tree(self, tree: Any) -> None:
        if isinstance(tree, dict):
            for key, tl_tree in tree.items():
                if key.upper() == "CHILDREN":
                    self.enumerate_object_tree(tree=tl_tree)
                elif key.upper() == "PROPERTIES":
                    self.enumerate_object_tree(tree=tl_tree)
                elif key.upper() == "ID":
                    self.id_list.append(tl_tree)
        elif isinstance(tree, list):
            for item in tree:
                self.enumerate_object_tree(tree=item)
        elif isinstance(tree, str):
            self.id_list.append(tree)
    
    def get_element(self) -> tuple[str, win32com.client.CDispatch | None]:
        try:
            self.element = self.session.findById(self.id)
            self.is_container = self.element.containerType
            self.name = self.element.name
            self.parent = self.element.parent
            self.type = self.element.type
            self.type_as_number = self.element.typeAsNumber
            return ("", self.element)
        except Exception as err:
            return (err, None)
    
    def visualize(self, delay: Optional[float] = 1.0) -> None:
        try:
            self.element.Visualize(True)
            time.sleep(delay)
            self.element.Visualize(False)
        except Exception as err:
            self.logger.log.error(f"Unhandled error during call to SapGuiFramework.SapGui.SapObject.SapObject.SapObject.visualize|{err}")
    
    def find_element(self, search: str) -> dict:
        result = process.extractOne(search, self.id_list, scorer=fuzz.partial_token_sort_ratio)
        return {"Search": search, "ID": result[0], "Score": result[1]}
    
    def get_element_history(self) -> list:
        if hasattr(self.element, "historyList"):
            return [i for i in self.element.historyList]
        return []
