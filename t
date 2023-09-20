Response Time (Hours) = 
VAR ResponseDateTime = RELATED('Sent Items'[DateTimeSent])
RETURN
IF(ISBLANK(ResponseDateTime), BLANK(), (ResponseDateTime - 'Inbox'[DateTimeReceived]) * 24)

Average Response Time (Hours) = 
AVERAGEX(
    FILTER('Inbox', NOT(ISBLANK('Sent Items'[DateTimeSent]))),
    'Inbox'[Response Time (Hours)]
)

