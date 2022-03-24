/*
 * Copyright (c) 2020 Automation Anywhere.
 * All rights reserved.
 *
 * This software is the proprietary information of Automation Anywhere.
 * You shall use it only in accordance with the terms of the license agreement
 * you entered into with Automation Anywhere.
 */

/**
 *
 */
package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.ListValue;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import static com.automationanywhere.commandsdk.model.AttributeType.BOOLEAN;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;
import static com.automationanywhere.commandsdk.model.DataType.LIST;
import static com.automationanywhere.commandsdk.model.DataType.STRING;

@BotCommand
@CommandPkg(label = "FileToBase64",
        description = "Esta action converte arquivos pdf para Base64", icon = "pkg.svg", name = "GetFilesFromPath",
        return_description = "", return_type = LIST,return_sub_type = STRING,  return_required = true)


public class GetFilesFromPath {

    @Execute
    public ListValue<String> action(
            @Idx(index = "1", type = TEXT)
            @Pkg(label = "Diretório")
            @NotEmpty String diretorio,
            @Idx(index = "2", type = TEXT)
            @Pkg(label = "Regex",default_value = ".*", default_value_type = STRING)
            String pattern,
            @Idx(index = "3", type = BOOLEAN)
            @Pkg(label = "Nome completo do arquivo:", description = "Nome completo do arquivo:",default_value = "", default_value_type = DataType.BOOLEAN)
            Boolean fullName
    ) {


        File folder = new File(diretorio);
        File[] listOfFiles = folder.listFiles();

        ListValue<String> returnvalue = new ListValue<String>();
        List<Value> files = new ArrayList<Value>();

        for (File file : listOfFiles) {
            if (file.isFile()) {
                if(file.getName().matches(pattern)) {
                    if(fullName)
                        try {
                            files.add(new StringValue(file.getCanonicalPath()));
                        }catch (IOException e) {
                            throw new BotCommandException(e.getMessage().toString());
                        }
                    else
                        files.add(new StringValue(file.getName()));
                }
            }
        }

        returnvalue.set(files);
        return returnvalue;


    }
    

}
