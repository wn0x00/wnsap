from win32com.client import Dispatch


class Application:
    def __init__(self):
        self.app = self._create_instance()

    def major_version(self):
        return self.app.MajorVersion

    def minor_version(self):
        return self.app.MinorVersion

    def connections(self):
        return self.app.Connections

    def active_session(self):
        return self.app.ActiveSession

    def open_connection(self, description, sync=True):
        self.app.OpenConnection(description, Sync=sync)

    def open_connection_by_connection_string(self, connect_string):
        self.app.OpenConnectionByConnectionString(connect_string)

    def _create_instance(self):
        return Dispatch("SapGui.Application")


if __name__ == "__main__":
    app = Application()

    pass
