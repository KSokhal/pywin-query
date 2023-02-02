from typing import List, Set

import pywintypes
from win32com.client import CDispatch, Dispatch

BASE_QUERY = "SELECT {', '.join(self.headers)} FROM SystemIndex WHERE SCOPE='file:{directory_path}'"


class WinQuery:
    """
    Class representing a Windows Search Query.
    Used to peform a query over files that are in Windows Search Index.
    """

    def __init__(
        self, directory_path: str, search_terms: List[str], requested_exts: Set[str], headers: List[str]
    ):

        #TODO: Check all the parameters are of the right type
        #TODO: Allow customer headers to be passed
        #TODO: Attempt to multi-thread the query running

        # https://msdn.microsoft.com/en-us/library/windows/desktop/bb419046(v=vs.85).aspx
        self.headers = [
            "System.ItemName",  # Name of file
            "System.ItemPathDisplay",  # Absolute path of file
            "System.ItemFolderPathDisplay",  # Directory path of file
            "System.FileExtension",  # File extension
        ]

        self.search_terms = search_terms
        self.requested_exts = requested_exts
        self.queries = []
        self.directory_path = directory_path
        self._construct_queries()

    def _construct_queries(self):
        """
        Construct a query for each of the search terms.
        Multiple queries are used so results can be assigned to a term
        """
        self.queries = []
        for term in self.search_terms:
            q = BASE_QUERY.format(", ".join(self.headers), self.directory_path) + f" and CONTAINS('{term}')"
            self.queries.append((term, q))

    def _get_connection(self) -> CDispatch:
        # Establish connection
        conn = Dispatch("ADODB.Connection")
        connstr = (
            "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
        )
        conn.Open(connstr)
        conn.CommandTimeout = 0  # remove timeout for searching


    def query(self, query, conn):
        results = []

        record_set = Dispatch("ADODB.Recordset")

        try:
            record_set.Open(query, conn)
        except Exception as e:
            raise RuntimeError(f"Failed to open query \n\t{query}\n\t: {e}")

        while not record_set.EOF:
            record_set.MoveNext()
            results.append(record_set.Fields)

        record_set.Close()
        return results


    def execute(self) -> List[List[str]]:

        results = []
        files_found = set()

        conn = self._get_connection()

        for term, query in self.queries:
            record_set = Dispatch("ADODB.Recordset")

            try:
                record_set.Open(query, conn)
            except Exception as e:
                raise RuntimeError(f"Failed to open query \n\t{query}\n\t: {e}")

            try:
                record_set.MoveFirst()
            except:
                # If recored set is EoF then no results were found and can continue
                # If it it is not then the query failed and all queries be cancelled
                if record_set.EOF:
                    record_set.Close()
                    record_set = None
                    continue
                else:
                    record_set.Close()
                    record_set = None
                    raise RuntimeError(f"Failed to move first in query \n\t{query}\n\t: {e}")

            while not record_set.EOF:
                abs_path = record_set.Fields.Item("System.ItemPathDisplay").Value

                if abs_path in files_found:
                    try:
                        record_set.MoveNext()
                    except pywintypes.com_error as e:
                        logger.error(
                            f"Failed to move next in query \n\t{query}\n\t: {e}"
                        )
                    continue
                else:
                    files_found.add(
                        abs_path,
                    )

                if (
                    record_set.Fields.Item("System.FileExtension").Value
                    in self.requested_exts
                ):
                    result_entry = [
                        term,
                        record_set.Fields.Item("System.ItemFolderPathDisplay").Value,
                        record_set.Fields.Item("System.ItemName").Value,
                    ]

                    results.append(result_entry)

                try:
                    record_set.MoveNext()
                except pywintypes.com_error as e:
                    logger.error(f"Failed to move next in query \n\t{query}\n\t: {e}")

            record_set.Close()
            record_set = None
        conn.Close()
        conn = None
        return results
