from win32com.client import Dispatch


class Application:
    def __init__(self):
        self.app = self._create_instance()

    @property.getter
    def major_version(self):
        self.app.MajorVersion

    @property.getter
    def minor_version(self):
        self.app.MinorVersion

    @property.getter
    def connections(self):
        self.app.Connections

    def active_session(self):
        pass

    def open_connection(self, description):
        self.app.OpenConnection(description)

    def open_connection_by_connection_string(self, connect_string):
        self.app.OpenConnectionByConnectionString(connect_string)

    def _create_instance(self):
        return Dispatch("SapGui.Application")
