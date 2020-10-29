from slacker import Slacker


class Slack:
    def __init__(self):
        self.token = 'xoxp-1445293978807-1472663966193-1445423374807-7ffe5066323c5c649e44c6adad984d29'

    def notification(self, pretext=None, title=None, fallback=None, text=None):
        attachments_dict = dict()
        attachments_dict['pretext'] = pretext   # test1
        attachments_dict['title'] = title       # test2
        attachments_dict['fallback'] = fallback # test3
        attachments_dict['text'] = text

        attachments = [attachments_dict]

        slack = Slacker(self.token)
        slack.chat.post_message(channel='#realmessage', text=None, attachments=attachments, as_user=None)