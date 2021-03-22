CAKE.COFFEE
¶
cake is a simplified version of Make (Rake, Jake) for CoffeeScript. You define tasks with names and descriptions in a Cakefile, and can call them from the command line, or invoke them from other tasks.

Running cake with no arguments will print out a list of all the tasks in the current directory's Cakefile.

External dependencies.

 
fs           = require 'fs'
path         = require 'path'
helpers      = require './helpers'
optparse     = require './optparse'
CoffeeScript = require './coffee-script'

existsSync   = fs.existsSync or path.existsSync
¶
Keep track of the list of defined tasks, the accepted options, and so on.

 
tasks     = {}
options   = {}
switches  = []
oparse    = null
¶
Mixin the top-level Cake functions for Cakefiles to use directly.

 
helpers.extend global,
¶
Define a Cake task with a short name, an optional sentence description, and the function to run as the action itself.

 
  task: (name, description, action) ->
    [action, description] = [description, action] unless action
    tasks[name] = {name, description, action}
¶
Define an option that the Cakefile accepts. The parsed options hash, containing all of the command-line options passed, will be made available as the first argument to the action.

 
  option: (letter, flag, description) ->
    switches.push [letter, flag, description]
¶
Invoke another task in the current Cakefile.

 
  invoke: (name, cb) ->
    missingTask name unless tasks[name]
    tasks[name].action options
¶
Run cake. Executes all of the tasks you pass, in order. Note that Node's asynchrony may cause tasks to execute in a different order than you'd expect. If no tasks are passed, print the help screen. Keep a reference to the original directory name, when running Cake tasks from subdirectories.

 
exports.run = (cb) ->
  global.__originalDirname = fs.realpathSync '.'
  process.chdir cakefileDirectory __originalDirname
  args = process.argv[2..]
  CoffeeScript.run fs.readFileSync('Cakefile').toString(), filename: 'Cakefile'
  oparse = new optparse.OptionParser switches
  return printTasks() unless args.length
  try
    options = oparse.parse(args)
  catch e
    return fatalError "#{e}"
  invoke arg for arg in options.arguments
¶
Display the list of Cake tasks in a format similar to rake -T

 
printTasks = ->
  relative = path.relative or path.resolve
  cakefilePath = path.join relative(__originalDirname, process.cwd()), 'Cakefile'
  console.log "#{cakefilePath} defines the following tasks:\n"
  for name, task of tasks
    spaces = 20 - name.length
    spaces = if spaces > 0 then Array(spaces + 1).join(' ') else ''
    desc   = if task.description then "# #{task.description}" else ''
    console.log "cake #{name}#{spaces} #{desc}"
  console.log oparse.help() if switches.length
¶
Print an error and exit when attempting to use an invalid task/option.

 
fatalError = (message) ->
  console.error message + '\n'
  console.log 'To see a list of all tasks/options, run "cake"'
  process.exit 1

missingTask = (task) -> fatalError "No such task: #{task}"
¶
When cake is invoked, search in the current and all parent directories to find the relevant Cakefile.

 
cakefileDirectory = (dir) ->
  return dir if existsSync path.join dir, 'Cakefile'
  parent = path.normalize path.join dir, '..'
  return cakefileDirectory parent unless parent is dir
  throw new Error "Cakefile not found in #{process.cwd()}"
