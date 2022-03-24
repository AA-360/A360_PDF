package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.Boolean;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class GetFilesFromPathCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(GetFilesFromPathCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    GetFilesFromPath command = new GetFilesFromPath();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("diretorio") && parameters.get("diretorio") != null && parameters.get("diretorio").get() != null) {
      convertedParameters.put("diretorio", parameters.get("diretorio").get());
      if(convertedParameters.get("diretorio") !=null && !(convertedParameters.get("diretorio") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","diretorio", "String", parameters.get("diretorio").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("diretorio") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","diretorio"));
    }

    if(parameters.containsKey("pattern") && parameters.get("pattern") != null && parameters.get("pattern").get() != null) {
      convertedParameters.put("pattern", parameters.get("pattern").get());
      if(convertedParameters.get("pattern") !=null && !(convertedParameters.get("pattern") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","pattern", "String", parameters.get("pattern").get().getClass().getSimpleName()));
      }
    }

    if(parameters.containsKey("fullName") && parameters.get("fullName") != null && parameters.get("fullName").get() != null) {
      convertedParameters.put("fullName", parameters.get("fullName").get());
      if(convertedParameters.get("fullName") !=null && !(convertedParameters.get("fullName") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","fullName", "Boolean", parameters.get("fullName").get().getClass().getSimpleName()));
      }
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("diretorio"),(String)convertedParameters.get("pattern"),(Boolean)convertedParameters.get("fullName")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
