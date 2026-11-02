#ifndef SAM_OFFER_GENERATOR_TEMPLATE_PROFILE_H
#define SAM_OFFER_GENERATOR_TEMPLATE_PROFILE_H

#include <map>
#include <string>

// Описание структуры конкретного Excel-шаблона.
struct TemplateProfile {
    // Уникальный идентификатор шаблона.
    std::string templateId;

    // Бренд или направление шаблона.
    std::string brand;

    // Сценарий расчёта шаблона.
    std::string scenario;

    // Название рабочего листа.
    std::string sheetName;

    // Карта смысловых полей и адресов ячеек Excel.
    std::map<std::string, std::string> cellMap;
};

#endif // SAM_OFFER_GENERATOR_TEMPLATE_PROFILE_H
