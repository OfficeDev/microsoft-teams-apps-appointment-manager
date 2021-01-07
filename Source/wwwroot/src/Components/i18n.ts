import i18nInstance, { i18n } from 'i18next';
import { initReactI18next } from 'react-i18next';
import Backend from 'i18next-http-backend';

export function i18nInit(fallbackLang: string): i18n {
    i18nInstance
        .use(Backend)
        .use(initReactI18next) // passes i18n down to react-i18next
        .init({
            fallbackLng: fallbackLang,
            ns: [],
            defaultNS: 'common',
            fallbackNS: 'common',
            keySeparator: false, // we do not use keys in form messages.welcome
            interpolation: {
                escapeValue: false, // react already safes from xss
            },
        });
    return i18nInstance;
}