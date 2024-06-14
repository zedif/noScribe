app_version = '0.5'

def ms_to_str(milliseconds: float, include_ms=False):
    """ Convert milliseconds into formatted timestamp 'hh:mm:ss' """
    seconds, milliseconds = divmod(milliseconds,1000)
    minutes, seconds = divmod(seconds, 60)
    hours, minutes = divmod(minutes, 60)
    formatted = f'{hours:02d}:{minutes:02d}:{seconds:02d}'
    if include_ms:
        formatted += f'.{milliseconds:03d}'
    return formatted

