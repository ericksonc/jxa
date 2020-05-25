(() => {const folderFilter =  {hidden: false}
const projectFilter = {effectiveStatus: "active status"}
const taskFilter =    t => {
    const completed = t.completed()
    const dropped = t.dropped()
    return !completed && !dropped
  }

const of = Application('OmniFocus')
const app = Application.currentApplication()
app.includeStandardAdditions = true
const dd = of.defaultDocument
const rootFolders = dd.folders
const rootTasks = dd.projects

let nodeText = `<?xml version="1.0" encoding="UTF-8"?>
<opml version="1.0">
<head>
<title>OmniFocus Export</title>
</head>
<body>`

const openNode = (name) => {
  name = name.replace(/"/g, '&quot;')
  name = name.replace(/&/g, '&amp;')
  name = name.replace(/</g, '&lt;')
  name = name.replace(/>/g, '&gt;')
  name = name.replace(/[^\x00-\x7F]/g, "");

  nodeText += `
  <outline text="${name}">`
}

const closeNode = () => {
  nodeText += `
  </outline>`}


const parseFolders = (fObj) => {
  const folders = fObj.whose(folderFilter)()
  folders.forEach(f => {
    openNode(f.name())
    if (f.folders.length) parseFolders(f.folders)
    if (f.projects.length) parseProjects(f.projects)
    closeNode()
  })
}

const parseProjects = (pObj) => {
  const projects = pObj()
  const tasks = projects.map(p => p.rootTask)
  parseTaskArray(tasks)
}

const parseTaskArray = (taskArray) => {
  const tasks = taskArray.filter(taskFilter)
  tasks.forEach(t => {
    openNode(t.name())
    if (t.tasks.length) parseTaskArray(t.tasks())
    closeNode()
  })
}

function writeTextToFile(text, file, overwriteExistingContent) {
    try {
 
        // Convert the file to a string
        var fileString = file.toString()
 
        // Open the file for writing
        var openedFile = app.openForAccess(Path(fileString), { writePermission: true })
 
        // Clear the file if content should be overwritten
        if (overwriteExistingContent) {
            app.setEof(openedFile, { to: 0 })
        }
 
        // Write the new content to the file
        app.write(text, { to: openedFile, startingAt: app.getEof(openedFile) })
 
        // Close the file
        app.closeAccess(openedFile)
 
        // Return a boolean indicating that writing was successful
        return true
    }
    catch(error) {
 
        try {
            // Close the file
            app.closeAccess(file)
        }
        catch(error) {
            // Report the error is closing failed
            console.log(`Couldn't close file: ${error}`)
        }
 
        // Return a boolean indicating that writing was successful
        return false
    }
}

parseFolders(rootFolders)
nodeText += `</body></opml>`
writeTextToFile(nodeText, app.pathTo("desktop").toString() + "/OmniFocusExport.opml", true)

})()