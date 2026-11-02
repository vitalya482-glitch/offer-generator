#ifndef SAM_OFFER_GENERATOR_CALC_FILE_CREATOR_H
#define SAM_OFFER_GENERATOR_CALC_FILE_CREATOR_H

#include <filesystem>
#include <string>

// Создаёт новый файл расчёта на основе выбранного шаблона.
class CalcFileCreator {
public:
    // Создаёт объект генератора файлов расчёта.
    CalcFileCreator();

    // Копирует шаблон в папку результата и возвращает путь нового файла.
    std::filesystem::path create(
        const std::filesystem::path& templatePath,
        const std::filesystem::path& outputDirectory,
        const std::string& fileName
    ) const;
};

#endif // SAM_OFFER_GENERATOR_CALC_FILE_CREATOR_H
