#ifndef SAM_OFFER_GENERATOR_CALC_FILE_FILLER_H
#define SAM_OFFER_GENERATOR_CALC_FILE_FILLER_H

#include "CalculationResult.h"
#include "TemplateProfile.h"

#include <filesystem>

// Заполняет созданный Excel-файл результатами расчёта.
class CalcFileFiller {
public:
    // Создаёт объект заполнения Excel-файлов.
    CalcFileFiller();

    // Записывает результат расчёта в файл по профилю шаблона.
    void fill(
        const std::filesystem::path& workbookPath,
        const TemplateProfile& profile,
        const CalculationResult& result
    ) const;
};

#endif // SAM_OFFER_GENERATOR_CALC_FILE_FILLER_H
