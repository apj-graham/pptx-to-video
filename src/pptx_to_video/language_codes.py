from deep_translator import GoogleTranslator


def language_to_code(language: str) -> str:
    """Convert language name to language code using deep_translator

    :param language: Language name (e.g., 'English', 'Spanish', 'French')
    :type language: str
    :return: Language code (e.g., 'en', 'es', 'fr')
    :rtype: str
    """
    translater = GoogleTranslator()
    langs = translater.get_supported_languages(as_dict=True)
    try:
        return langs[language.lower()]
    except KeyError:
        raise ValueError(f"Language '{language}' not supported.")
