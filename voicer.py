import win32com.client


def voice(s):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(s)
