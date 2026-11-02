#include "ExcelTemplateRecognizer.h"

#include <stdexcept>

// Создаёт распознаватель Excel-шаблонов.
ExcelTemplateRecognizer::ExcelTemplateRecognizer() = default;

// Распознаёт шаблон по указанному файлу.
TemplateProfile ExcelTemplateRecognizer::recognize(const std::filesystem::path& templatePath) const {
    if (templatePath.empty()) {
        throw std::invalid_argument("Путь к Excel-шаблону не указан");
    }

    if (!std::filesystem::exists(templatePath)) {
        throw std::runtime_error("Excel-шаблон не найден: " + templatePath.string());
    }

    // Полное чтение OOXML-шаблона подключим отдельным этапом.
    TemplateProfile profile;
    profile.templateId = templatePath.stem().string();
    return profile;
}
