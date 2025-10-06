using DocumentProcessingLibrary.Core.Interfaces;

namespace DocumentProcessingLibrary.Core.Strategies.Search;

/// <summary>
    /// Предопределенные стратегии для типовых задач
    /// </summary>
    public static class CommonSearchStrategies
    {
        /// <summary>
        /// Поиск десятичных обозначений изделий
        /// </summary>
        public static ITextSearchStrategy DecimalDesignations =>
            new RegexSearchStrategy(
                "DecimalDesignations",
                new RegexPattern(
                    "FullDesignation",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+\.(?:[0-9]{2,2}\.){2,}[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?"
                ),
                new RegexPattern(
                    "ShortDesignation",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9-]+-[А-Я0-9-]+\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                ),
                new RegexPattern(
                    "MinimalDesignation",
                    @"(?=[А-Я0-9-]*[А-Я])[А-Я0-9]+\.[0-9]{2,2}\.[0-9]{3,3}(?:ТУ)?[.,;:!?\-]?\b"
                )
            );

        /// <summary>
        /// Поиск имен в формате "Фамилия И. О." и "И. О. Фамилия"
        /// </summary>
        public static ITextSearchStrategy PersonNames =>
            new RegexSearchStrategy(
                "PersonNames",
                new RegexPattern(
                    "SurnameFirst",
                    @"[А-Я][а-я]+\s[А-Я]\.\s?[А-Я]\."
                ),
                new RegexPattern(
                    "InitialsFirst",
                    @"[А-Я]\.\s?[А-Я]\.\s?[А-Я][а-я]+"
                )
            );

        /// <summary>
        /// Поиск email адресов
        /// </summary>
        public static ITextSearchStrategy EmailAddresses =>
            new RegexSearchStrategy(
                "EmailAddresses",
                new RegexPattern(
                    "Email",
                    @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
                )
            );

        /// <summary>
        /// Поиск телефонных номеров (российский формат)
        /// </summary>
        public static ITextSearchStrategy PhoneNumbers =>
            new RegexSearchStrategy(
                "PhoneNumbers",
                new RegexPattern(
                    "RussianPhone",
                    @"(\+7|8)[\s-]?\(?[0-9]{3}\)?[\s-]?[0-9]{3}[\s-]?[0-9]{2}[\s-]?[0-9]{2}"
                )
            );
    }