#ifndef SAM_OFFER_GENERATOR_POSITION_PARSER_H
#define SAM_OFFER_GENERATOR_POSITION_PARSER_H

#include "Position.h"

#include <filesystem>
#include <vector>

// Базовый интерфейс для получения позиций из внешнего источника.
class PositionParser {
public:
    // Освобождает ресурсы базового парсера.
    virtual ~PositionParser();

    // Читает позиции из указанного источника.
    virtual std::vector<Position> parse(const std::filesystem::path& sourcePath) const = 0;
};

#endif // SAM_OFFER_GENERATOR_POSITION_PARSER_H
