//
// Created by Troldal on 05/07/2023.
//

#ifndef XLTHERMO_DATA_DATABASE_HPP
#define XLTHERMO_DATA_DATABASE_HPP

#include <string>
#include <vector>

class Database final
{
public:
    static Database& instance()
    {
        static Database db;    // The one, unique instance
        return db;
    }

    std::vector<std::string> listNames() const;

    std::vector<std::string> listCAS() const;

    std::vector<std::string> listComponents() const;







private:
    Database() {}                                     // NOLINT
    Database(Database const&)            = delete;    // NOLINT
    Database& operator=(Database const&) = delete;    // NOLINT
    Database(Database&&)                 = delete;    // NOLINT
    Database& operator=(Database&&)      = delete;    // NOLINT
    // ... Potentially some data members
};

#endif    // XLTHERMO_DATA_DATABASE_HPP
