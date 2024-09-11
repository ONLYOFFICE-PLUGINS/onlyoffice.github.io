var library = {};

library["Current Word"] = "(async () => {\n\
    return await AI.callMethod(\"GetCurrentWord\");\n\
})(entryPoint);";

library["Synonyms"] = "(async () => {\n\
    if (!entryPoint)\n\
        return \"[]\";\n\
    return await AI.provider(\"gpt-3.5\").request(`Give synonyms for the word \"${entryPoint}\" as javascript array`);\n\
})(entryPoint);";

library["AddComment"] = "(async () => {\n\
    return await AI.callMethod(\"AddComment\", [{\n\
        Username : \"AI\",\n\
        Text : \"Synonyms: \" + entryPoint,\n\
        Time: Date.now(),\n\
        Solver: false\n\
    }]);\n\
})(entryPoint);";

library["Selection As Text"] = "(async () => {\n\
    return await AI.callMethod(\"GetSelectedText\");\n\
})(entryPoint);";

library["Insert As Text"] = "(async () => {\n\
    Asc.scope.data = (entryPoint || \"\").split('\n\n');\n\
    await AI.callCommand(function() {\n\
        let oDocument = Api.GetDocument();\n\
        for (let ind = 0; ind < Asc.scope.data.length; ind++) {\n\
            let text = Asc.scope.data[ind];\n\
            if (text.length) {\n\
                let oParagraph = Api.CreateParagraph();\n\
                oParagraph.AddText(text);\n\
                oDocument.Push(oParagraph);\n\
            }\n\
        }\n\
    });\n\
})(entryPoint);";

library["Summarize"] = "(async () => {\n\
    let text = entryPoint || \"\";\n\
    return await AI.provider(\"gpt-3.5\").request(`Summarize this text: \"${text}\"`);\n\
})(entryPoint);";