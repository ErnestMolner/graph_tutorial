from graph_utils import authentication, outlook_utils

graph_app = authentication.Application('Tenant ID', 'APP ID', 'Credential secret name')
outlook_utils.send_mail_with_attatchments(graph_app, ['Picture1.png', 'test.txt'], "Hello how are you doing?", "They were <b>awesome</b>!", "test@test.com", ["user1@test.com"], ["bccrecipient@test.com"])