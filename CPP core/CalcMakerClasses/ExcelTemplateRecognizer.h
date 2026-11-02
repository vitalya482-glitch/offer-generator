#ifndef SAM_OFFER_GENERATOR_EXCEL_TEMPLATE_RECOGNIZER_H
#define SAM_OFFER_GENERATOR_EXCEL_TEMPLATE_RECOGNIZER_H

#include "TemplateProfile.h"

#include <filesystem>

// Распознаёт тип и структуру выбранного Excel-шаблона.
class ExcelTemplateRecognizer {
public:
    // Создаёт распознаватель Excel-шаблонов.
    ExcelTemplateRecognizer();

    // Распознаёт шаблон по указанному файлу.
    TemplateProfile recognize(const std::filesystem::path& templatePath) const;
};

#endif // SAM_OFFER_GENERATOR_EXCEL_TEMPLATE_RECOGNIZER_H
