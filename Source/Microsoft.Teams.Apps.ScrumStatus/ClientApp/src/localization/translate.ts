
const defaultLocale = 'en-US';

export function getUserLocale() {
    if (navigator.languages) {
        return navigator.languages[0].toLowerCase();
    }
    else if (navigator.language) {
        return navigator.language.toLowerCase();
    }
    else {
        return defaultLocale;
    }
}