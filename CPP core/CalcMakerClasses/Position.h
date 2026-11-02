#ifndef SAM_OFFER_GENERATOR_POSITION_H
#define SAM_OFFER_GENERATOR_POSITION_H

#include <string>

// Одна товарная позиция, подготовленная для расчёта.
struct Position {
    // Уникальный или исходный код позиции.
    std::string code;

    // Наименование или модель позиции.
    std::string name;

    // Бренд позиции.
    std::string brand;

    // Количество позиции.
    double quantity = 0.0;

    // Цена одной единицы для расчёта.
    double unitPrice = 0.0;

    // Общая цена позиции, если она уже известна из источника.
    double totalPrice = 0.0;

    // Валюта цены позиции.
    std::string currency;

    // Источник, из которого была получена позиция.
    std::string source;

    // Дополнительное описание позиции или её комплектации.
    std::string description;
};

#endif // SAM_OFFER_GENERATOR_POSITION_H
