import unittest
import authentication
import outlook_utils


# The tests should run using the graph_utils directory as test root.
# The self.appSecretReference within the config file needs to be a generic windows credential name. With the password being the application secret.
class Test_outlook(unittest.TestCase):

    def setUp(self):
        self.authority = "Tenant ID"
        self.client_id = "APP ID"
        self.appSecretReference = "Credential secret name"
        self.graph_app = authentication.Application(self.authority, self.client_id, self.appSecretReference)

    def test_authentication(self):
        autentication = self.graph_app.acquire_token()
        self.failIfEqual(None, autentication)

    def test_send_mail_with_attatchments(self):
        outlook_utils.send_mail_with_attatchments(self.graph_app, ['graph_utils/test/Picture1.png', 'graph_utils/test/test.txt'], "Hello how are you doing?", "They were <b>awesome</b>!", "test@test.com", ["user1@test.com"], ["bccrecipient@test.com"])

    def test_send_mail(self):
        outlook_utils.send_mail(self.graph_app, "test subject", "nice body", "test@test.com", ["user1@test.com"], ["bccrecipient@test.com"])

    def test_read_mail(self):
        outlook_utils.read_mail_from_folder(self.graph_app, "test@test.com", "Inbox")
