#include "CalcFileFiller.h"

#include <stdexcept>

// Создаёт объект заполнения Excel-файлов.
CalcFileFiller::CalcFileFiller() = default;

// Записывает результат расчёта в файл по профилю шаблона.
void CalcFileFiller::fill(
    const std::filesystem::path& workbookPath,
    const TemplateProfile& profile,
    const CalculationResult& result
) const {
    if (workbookPath.empty() || !std::filesystem::exists(workbookPath)) {
        throw std::runtime_error("Файл расчёта для заполнения не найден");
    }

    if (profile.templateId.empty()) {
        throw std::invalid_argument("Профиль Excel-шаблона не определён");
    }

    // Заполнение OOXML-файла подключим после выбора библиотеки Excel.
    (void)result;
}
