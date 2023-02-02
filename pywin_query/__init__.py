from typing import Dict, List, Union

import pywintypes
from win32com.client import CDispatch, Dispatch

BASE_QUERY = "SELECT {headers} FROM SystemIndex WHERE SCOPE='file:{directory_path}'"


class WinQuery:
    """
    Class representing a Windows Search Query.
    Used to peform a query over files that are in Windows Search Index.
    """

    def __init__(
        self,
        directory_path: str,
        headers: List[str] = [
            "System.ItemPathDisplay",  # Absolute path of file
        ],
    ):
        self.directory_path = directory_path
        assert isinstance(self.directory_path, str)

        # More headers, such as file extentions and dirextory paths, can be found at the link below
        # https://msdn.microsoft.com/en-us/library/windows/desktop/bb419046(v=vs.85).aspx
        self.headers = headers
        assert isinstance(self.headers, list)

    def _construct_query(self, search_term: Union[str, List[str]]) -> str:
        """Construct a query for a given search term(s)."""
        base_query = BASE_QUERY.format(
            headers=", ".join(self.headers), directory_path=self.directory_path
        )

        if isinstance(search_term, str):
            return base_query + f" and CONTAINS('{search_term}')"
        else:
            q = base_query + f" and ("
            for index, value in enumerate(search_term):
                q += f"CONTAINS('{value}')"
                if index == len(search_term) - 1:
                    q += ")"
                    break
                q += " or "
        return q

    def _get_connection(self) -> CDispatch:
        """Get connection to Windows Internal Database."""

        # Establish connection
        conn = Dispatch("ADODB.Connection")
        connstr = (
            "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
        )
        conn.Open(connstr)
        conn.CommandTimeout = 0  # remove timeout for searching
        return conn

    def exc_query(self, query_string, conn):
        """Execute a single query."""
        results = []

        record_set = Dispatch("ADODB.Recordset")

        try:
            record_set.Open(query_string, conn)
        except Exception as e:
            raise RuntimeError(f"Failed to open query \n\t{query_string}\n\t: {e}")

        while not record_set.EOF:

            if len(self.headers) == 0:
                raise ValueError("No headers provided")

            cur_res = []

            for h in self.headers:
                try:
                    cur_res.append(record_set.Fields.Item(h).Value)
                except Exception as e:
                    raise ValueError("Headers {h} does not exist on record set.")

            record_set.MoveNext()

            cur_res = cur_res[0] if len(cur_res) == 1 else cur_res
            results.append(cur_res)

        record_set.Close()
        return results

    def query(
        self, search_terms: Union[str, List[str]]
    ) -> Union[List[str], Dict[str, List[str]]]:
        """Execute a query for term(s)."""

        if not isinstance(search_terms, (str, list)):
            raise TypeError("Search terms must be a string for list of strings.")

        if isinstance(search_terms, list):
            for st in search_terms:
                if not isinstance(st, str):
                    raise TypeError("All values in search terms must be strings.")

        q = self._construct_query(search_terms)
        conn = self._get_connection()
        return self.exc_query(q, conn)
