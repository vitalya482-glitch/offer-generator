#ifndef SAM_OFFER_GENERATOR_CALC_MAKER_H
#define SAM_OFFER_GENERATOR_CALC_MAKER_H

#include "CalculationRequest.h"
#include "CalculationResult.h"

// Центральный класс расчётного ядра коммерческих предложений.
class CalcMaker {
public:
    // Создаёт объект расчётного ядра.
    CalcMaker();

    // Освобождает ресурсы объекта расчётного ядра.
    ~CalcMaker();

    // Выполняет расчёт по входному запросу.
    CalculationResult calculate(const CalculationRequest& request) const;

    // Проверяет входные данные расчёта.
    void validate(const CalculationRequest& request) const;

    // Сбрасывает внутреннее состояние расчётного ядра.
    void reset();
};

#endif // SAM_OFFER_GENERATOR_CALC_MAKER_H
