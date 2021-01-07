# Adding support for a new locale

The app template supports localization for the bot, frontend components (tabs and task modules), and Teams manifests. By default, the app template supports the `en-US` locale. You can add support for more locales by following the guidance below.

## Localize the bot

### Translate strings

1. Create a copy of the default English resource file, `Source\Resources\SharedResources.resx`, and name it `Source\Resources\SharedResources.{language-code}.resx`.
    - For example, if you want to support Spanish, name the new file `Source\Resources\SharedResources.es.resx`

1. Open the new resource file and replace the string **values** with strings in the desired language. Don't change the string **names**.

### Add locale to supported locales

The list of supported locales is stored as a comma-delimited application setting in the App Service. When deploying the app using the provided ARM template, this setting comes from the `supportedCultures` template parameter.

To add the new locale to the list of supported locales, include the locale in this parameter when deploying the app. For example, to add support for Spanish, set `supportedCultures` to `en-US,es`. Alternatively, if you've already deployed the app, you can update the app settings through the Azure portal or Azure CLI (see the [docs on configuring an App Service app](https://docs.microsoft.com/en-us/azure/app-service/configure-common)).

> Note: Even if you update the app setting for an existing deployment, you will still have to redeploy the code for the new `.resx` file to be included.

## Localize the frontend components

### Translate strings

1. Create a copy of the `en-US` locale folder, `Source\wwwroot\locales\en-US`, and name it `Source\wwwroot\locales\{locale}`.
    - For example, if you want to support Spanish generally, name the new folder `Source\wwwroot\locales\es`.
    - If you want to support a specific locale, like "Spanish - Spain", you can name the new folder accordingly: `Source\wwwroot\locales\es-ES`.

1. For each of the JSON files in the new locale directory, open the JSON file and replace the string **values** with strings in the desired language. Don't change the string **keys**.

### Add locale to supported locales

1. Open the app's webpack config at `Source\webpack\webpack.common.js`.

1. Find the `supportedLocales` array near the top of the file. Add the new locale to this array.
    - For example, if you want to support Spanish generally, the line would be
        ```js
        const supportedLocales = ['en', 'es'];
        ```

> Note: The frontend React project is bundled using webpack. By default, the generated bundle would include *all* locales from JS libraries like `moment.js` and `date-fns`. To avoid inflating the app's bundle, the included webpack configuration prevents bundling these locales, except the ones explicitly included in `supportedLocales` above.<br><br>If you prefer including all locales, remove the two `ContextReplacementPlugin`s from `webpack.common.js`. This will eliminate the need to add locales to `supportedLocales`.

## Localize the Teams manifests

Other app strings shown by Teams, such as tab names and app metadata, are defined in the `manifest.json` files, so these files need to be localized too. For more details, see the [doc on localizing your app manifest](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/apps-localization#localizing-the-strings-in-your-app-manifest).

The steps below are written for the staff member app, but you should repeat the steps for the admin app.

### Translate strings

1. Create a new manifest localization file: `Manifest\AgentApp\{language-code}.json`.
    - For example, if you want to support Spanish, name the file `Manifest\AgentApp\es.json`.

1. Open the new file in a text editor and paste the following JSON:

    ```json
    {
        "$schema": "https://developer.microsoft.com/json-schemas/teams/v1.7/MicrosoftTeams.schema.json",
        "name.short": "",
        "name.full": "",
        "description.short": "",
        "description.full": "",
        "staticTabs[0].name": ""
    }
    ```

1. Fill in the string values in the JSON above. For reference, see the corresponding English values in `Manifest\AgentApp\manifest.json`.

### Add locale to supported locales

1. Open the main `manifest.json` in a text editor: `Manifest\AgentApp\manifest.json`.

1. Find the `localizationInfo.additionalLanguages` array. Add a new object to the array for your new language. For example, if you want to support Spanish, add:

    ```json
    {
        "languageTag": "es",
        "file": "es.json"
    }
    ```

# Changing the default locale

The default locale is used in cases where a specific user's locale cannot be used. For example:

- Teams have multiple staff members who may have different locale preferences. Therefore, messages posted to channels are always sent in the default locale.
- A user may be using a locale that is not supported by the app. In this case, the app will fall back to the default locale.
- When the app is first installed by a staff member, the staff member's locale is not yet known. Therefore, the 1:1 welcome message to the staff member is always sent in the default locale.

The default locale for the app is `en-US`, but this can be changed by following the steps below.

## Change the bot default locale

### Rename resource files

For the backend resource files, the default resource file in `Source\Resources\` is named without a specific locale in the filename: `SharedResources.resx`. The resource file for the new default locale needs to have this name instead.

1. Rename the current default resource file from `SharedResources.resx` to `SharedResources.en.resx`.

1. Rename the new default locale's resource file from `SharedResources.{language-code}.resx` to `SharedResources.resx`.

### Change app setting

The default locale is stored as an application setting in the App Service. When deploying the app using the provided ARM template, this setting comes from the `defaultCulture` template parameter.

To switch to a different default locale, set this parameter when deploying the app. Alternatively, if you've already deployed the app, you can update the app settings through the Azure portal or Azure CLI. For more details, see the [docs on configuring an App Service app](https://docs.microsoft.com/en-us/azure/app-service/configure-common).

> Note: Even if you update the app setting for an existing deployment, you will still have to redeploy the code for the new `.resx` file to be included.

## Change the frontend default locale

The frontend retrieves the default locale from the backend, so no change is needed in the frontend code. Browsers may cache the default locale for the frontend, so users may not see the change immediately.

## Change the Teams manifest default locale

The steps below are written for the staff member app, but you should repeat the steps for the admin app.

### Move English strings to localization file (optional)

If you want to continue supporting English as a non-default language, you should move the existing English strings in `Manifest\AgentApp\manifest.json` to a separate English localization file. Follow the steps above for [localizing the Teams manifests](#localize-the-teams-manifests) to create the English file and add English as a supported language.

### Update manifest.json

The strings in `manifest.json` are used as the default, so these strings need to be changed from English to the new default language.

1. Open `Manifest\AgentApp\manifest.json` in a text editor.

1. Update the values of the following properties with strings in the new default language:
    - `name.short`
    - `name.full`
    - `description.short`
    - `description.full`
    - `staticTabs[0].name`

1. Update `localizationInfo.defaultLanguageTag` to the new default language.

### Delete old localization file (optional)

If you already had an existing manifest localization file for the new default language, you can safely delete it. The new default language strings are now in `manifest.json`.

1. Delete `Manifest\AgentApp\{language-code}.json`.

1. Remove the language from the `localizationInfo.additionalLanguages` array.