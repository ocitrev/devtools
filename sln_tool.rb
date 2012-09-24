#
# Outil pour faire des operations sur des fichiers solutions .sln
#

def verbose(text)
  puts "Verbose: #{text}" if $verbose
end

class Project
  
  attr_accessor :name
  attr_accessor :path
  attr_accessor :guid
  attr_accessor :data
  attr_accessor :parent
  attr_accessor :type
  attr_accessor :depends

  GUID_FOLDER = '2150e333-8fdc-42a3-9474-1a3956d46de8'
  GUID_CSHARP = 'fae04ec0-301f-11d3-bf4b-00c04f79efbc'
  GUID_CPP = '8bc9ceb8-8b4a-11d0-8d11-00a0c91bc942'

  include Comparable

  def initialize(name, path, type, guid)
    @parent = nil
    @data = []
    @name = name
    @path = path
    @type = type
    @guid = guid
    @depends = []
  end
  
  def is_folder?
    @type == GUID_FOLDER
  end

  def is_cpp?
    @type == GUID_CPP
  end

  def is_csharp?
    @type == GUID_CSHARP
  end
  
  def to_s
    name
  end
  
  def <=>(other)
    if is_folder? == other.is_folder?
      
      return -1 if parent.nil? and !other.parent.nil?
      return 1 if !parent.nil? and other.parent.nil?
      
      cmp = (parent <=> other.parent)
      if (cmp == 0)
        cmp = name.downcase <=> other.name.downcase
        if (cmp == 0)
          cmp = path.downcase <=> other.path.downcase
          if (cmp == 0)
            cmp = guid.downcase <=> other.guid.downcase
          end
        end
      end
      return cmp
    else
      if is_folder?
        return -1
      else
        return 1
      end
    end
  end
  
  def ==(other)
    
    case other
    when Project
      return @guid == other.guid
    when String
      return true if name.downcase == other.downcase
      return true if guid == other.downcase
    end
    
    return false
  end
  
  def eql?(other)
    return @guid == other.guid if other.class == Project
    return false
  end
  
end

class Solution

  def initialize(filepath)
    @depends_loaded = false
    parse(filepath)
  end

  private
  
  def parse(input)
    
    verbose "parsing #{input} ..."
    
    @lines = File.open(input, 'r') { |f| f.readlines }
    @slnpath = File.dirname(File.expand_path(input))
    @project_list = Array.new
    @projectguid_map = Hash.new
    @projectname_map = Hash.new { |h,k| h[k] = Array.new }
    @file_start = Array.new
    @file_end = Array.new
    @global = Array.new
    projet = nil
    inglobal = false
    nestedproject = false
    
    @lines.map do |l|
      
      if projet.nil?
        # est-ce que c'est une ligne de projet?
        md = /^\s*Project\("{([-a-fA-F0-9]+)}"\)\s*=\s*"(.*?)"\s*,\s*"(.*?)"\s*,\s*"{([-a-fA-F0-9]+)}"/.match(l)
        
        if md
          guid = md[4].downcase
          name = md[2]
          projet = Project.new(name, md[3], md[1].downcase, guid)
          raise "Project GUID {#{guid}} is already in map" unless @projectguid_map[guid].nil?
          @projectguid_map[guid] = projet
          # ajoute le projet dans une liste; la relation est 1 a n
          @projectname_map[name.downcase] << projet
          @project_list << projet
        end
      end
      
      if projet
        
        # ajoute les lignes du projet dans l'array
        projet.data << l
        
        # est-ce la fin du projet courrant ?
        projet = nil if l =~ /^\s*EndProject\s*$/
        
      else
        
        inglobal = (l =~ /^\s*Global\s*$/) unless inglobal
        
        if inglobal
          
          # la section NestedProjects determine la hierarchie des projets
          if nestedproject
            nestedproject = l !~ /^\s*EndGlobalSection\s*$/
            
            if nestedproject
              # assigne le parent du projet
              md = /{([-a-fA-F0-9]+)}\s*=\s*{([-a-fA-F0-9]+)}/.match(l)
              @projectguid_map[md[1].downcase].parent = @projectguid_map[md[2].downcase]
            end
            
          else
            nestedproject = l =~ /^\s*GlobalSection\(NestedProjects\)/
            inglobal = l !~ /^\s*EndGlobal\s*$/
          end
          
          @global << l
          
        else
          if @global.count > 0
            @file_end << l
          else
            @file_start << l
          end
        end
        
      end
    end
  end
  
  def load_dependencies
    
    verbose "loading dependencies ..."
    
    require 'rexml/document'
    project_list = @project_list.reject {|p| p.is_folder? }
    
    # trouve les dependances dans le fichier de solution
    project_list.each do |p|
      md = p.data.join.match(/ProjectSection\(ProjectDependencies\) = postProject(.+?)EndProjectSection/m)
      md.to_s.scan(/{([-a-fA-F0-9]+)} = {\1}/).each do |ref|
        p.depends << @projectguid_map[ref[0].downcase]
      end unless md.nil?
    end
    
    # trouve les dependances dans les fichiers de projets
    project_list.each do |p|
      path = File.join(@slnpath, p.path)
      verbose "parsing #{File.basename(p.path)} ..."
      File.open(path, 'r') do |f|
        doc = REXML::Document.new(f)
        doc.root.each_element('/Project/ItemGroup/ProjectReference') do |ref|
          guidRef = ref.text('Project').gsub(/^{|}$/, '')
          nameRef = ref.text('Name')
          dep = @projectguid_map[guidRef.downcase]
          dep = find_project(nameRef) if dep.nil? && !nameRef.nil?
          p.depends << dep unless dep.nil?
        end
      end
    end
    
    @depends_loaded = true
    
  end

  public
  
  def sort
    @project_list.sort!
  end
  
  def save(filepath)
    File.open(filepath, 'w') do |f|
      f.puts @file_start unless @file_start.empty?
      
      @project_list.each do |p|
        f.puts p.data unless p.data.empty?
      end
      
      f.puts @global unless @global.empty?
      f.puts @file_end unless @file_end.empty?
    end
  end
  
  def find_project(project)
    key = project.downcase
    p = @projectguid_map[key]
    
    if p.nil?
      lst = @projectname_map[key]
      
      if lst.size > 0
        if (lst.size == 1)
          p = lst[0]
        else
          raise "Project name '#{project}' is ambigous."
        end
      end
    end
    
    return p
  end
  
  def get_using(project)
    
    case project
    
    when String
      p = find_project(project)
      raise "Cannot find project '#{project}'" if p.nil?
      get_using(p)
      
    when Project
      
      load_dependencies unless @depends_loaded
      @project_list.select do |p|
        p.depends.count(project) > 0
      end
      
    end
    
  end
  
  def get_dependencies(project=nil, parent_list=nil)
    
    load_dependencies unless @depends_loaded
    
    case project
    
    when nil
      # pas de parametre, liste toutes les dependances
      list = []
      verbose 'building dependency tree'
      @project_list.reject{ |p| p.is_folder? }.each do |p|
        list |= get_dependencies(p, list) if (parent_list.nil? || parent_list.count(p) == 0)
      end
      return list
      
    when String
      # j'ai recu une string trouve un projet qui pourrait correspondre
      p = find_project(project)
      raise "Cannot find project '#{project}'" if p.nil?
      return get_dependencies(p)
      
    when Project
      list = []
      # optimisation: verification si le projet n'est pas dans la liste du parent
      if (parent_list.nil? || parent_list.count(project) == 0)
        project.depends.each do |p|
          list |= get_dependencies(p, list) if (parent_list.nil? || parent_list.count(p) == 0)
        end
        
        # ajoute le projet a la liste
        list << project
      end
      return list
      
    end
    
  end

end

require 'ostruct'

class SlnTool
  
  def get_command(args, options)
    ok = false
  
    case args[0]
    when /^sort$/i
      options.command = 'sort' unless options.nil?
      text = <<EOS
#{File.basename(__FILE__)} SORT solution_file [OUTPUT_FILE]

sorts the solution file
    This command will sort by alphabetical order all the projects of a solution file.
    The file is then saved.

options
    [OUTPUT_FILE]             Specifies an alternate output file
    
EOS
      if !options.nil? && args.size > 1
        options.input_file = args[1]
        options.output_file = args[1]
        options.output_file = args[2] if args.size > 2
        ok = true
      end
      
    when /^using$/i
      options.command = 'using' unless options.nil?
      text = <<EOS
#{File.basename(__FILE__)} USING solution_file PROJECT

List the projects that have 'PROJECT' as a dependency
    This command will load the solution file and find projects that have the specified
    project as a dependency.

options
    PROJECT                   The project name or project GUID
    
EOS
      if !options.nil? && args.size > 2
        options.input_file = args[1]
        options.project = args[2]
        ok = true
      end
      
    when /^depends$/i
      options.command = 'depends' unless options.nil?
      text = <<EOS
#{File.basename(__FILE__)} DEPENDS solution_file [PROJECT]

List the dependencies of the specified project
    This command will load the solution file and find all dependencies for the specified
    project.  If a project is not specified all dependencies are listed.
    
options
    [PROJECT]                 The project name or project GUID, optional.
    
EOS
      if !options.nil? && args.size > 1
        options.input_file = args[1]
        options.project = args[2]
        ok = true
      end
      
    when /^help$/i
      options.command = 'help' unless options.nil?
      if args.size > 1
        ok, text = get_command(args[1..-1], nil)
      else
        text = <<EOS
#{File.basename(__FILE__)} HELP [comamnd]

command
    SORT          Sort and save the solution file
    USING         List the projects that have 'project' as a dependency
    DEPENDS       List the dependencies of the specified project
    HELP          Shows help about commands
    
EOS
        ok = true
      end
    else
      options.command = nil unless options.nil?
      text = <<EOS
Usage: #{File.basename(__FILE__)} [ SORT | USING | DEPENDS | HELP ]

EOS
    end
    
    [ok, text]
  end

  def parse(args)
    options = OpenStruct.new
    ok, help_text = get_command(args, options)
   
    # la liste est vide, affiche le message d'aide
    if (!ok || options.command == 'help')
      puts help_text
      exit
    end
    
    options
  end
  
  def command_sort(options)
    sln = Solution.new(options.input_file)
    sln.sort
    sln.save options.output_file
  end
  
  def command_using(options)
    sln = Solution.new(options.input_file)
    puts sln.get_using(options.project).map { |p| p.name }
    puts
  end

  def command_depends(options)
    sln = Solution.new(options.input_file)
    puts sln.get_dependencies(options.project).map { |p| p.name }
    puts
  end

  def self.execute(args)
    @@tool ||= SlnTool.new
    options = @@tool.parse(args)
    @@tool.send("command_#{options.command}".to_sym, options)
  end
  
end

SlnTool.execute(ARGV)
